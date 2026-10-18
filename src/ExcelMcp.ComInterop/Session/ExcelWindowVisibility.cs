using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal readonly record struct ExcelVisibilityResult(bool IsVisible, bool IsForeground);

internal interface IExcelBatchWindowControl
{
    ExcelVisibilityResult SetVisibility(bool visible);
    string? DialogVisibilityError { get; }
    string? DialogDetails { get; }
}

internal static class ExcelWindowVisibility
{
    internal static ExcelVisibilityResult Set(
        ExcelProcessIdentity identity, bool visible, ILogger logger)
    {
        using var process = Process.GetProcessById(identity.ProcessId);
        if (process.HasExited || process.StartTime.ToUniversalTime().ToFileTimeUtc() != identity.StartedAtUtcFileTime)
            throw new InvalidOperationException("The session's Excel process identity could not be confirmed.");

        var inspection = ExcelDialogProbe.Inspect(identity, logger);
        if (inspection.State == WorkbookRefreshState.Unknown)
            throw new InvalidOperationException("Excel's window ownership could not be confirmed.");
        var roots = inspection.Windows
            .Where(window => window.ProcessId == identity.ProcessId && window.ClassName == "XLMAIN")
            .ToArray();
        foreach (var root in roots)
        {
            if (process.HasExited || GetWindowThreadProcessId(root.Handle, out var pid) == 0 || pid != identity.ProcessId)
                throw new InvalidOperationException("Excel window ownership changed before its visibility could be updated.");
            if (root.Visible == visible && (!visible || !IsIconic(root.Handle)))
                continue;
            if (!ShowWindowAsync(root.Handle, visible ? IsIconic(root.Handle) ? 9 : 5 : 0))
                throw new Win32Exception(Marshal.GetLastWin32Error(), "Could not request Excel window visibility.");
        }

        if (!SpinWait.SpinUntil(() =>
            {
                inspection = ExcelDialogProbe.Inspect(identity, logger);
                var currentRoots = inspection.Windows
                    .Where(window => window.ProcessId == identity.ProcessId && window.ClassName == "XLMAIN")
                    .ToArray();
                if (inspection.State == WorkbookRefreshState.Unknown || currentRoots.Length != roots.Length ||
                    currentRoots.Any(current => !roots.Any(original => original.Handle == current.Handle)))
                    throw new InvalidOperationException(
                        "Excel window ownership changed during the visibility request. Inspect Excel and retry.");
                return currentRoots.All(window => window.Visible == visible);
            }, TimeSpan.FromSeconds(1)))
            throw new InvalidOperationException("Excel did not confirm the requested window visibility within 1 second.");
        if (process.HasExited)
            throw new InvalidOperationException("Excel exited while its window visibility was changing.");

        var foreground = nint.Zero;
        if (visible)
        {
            foreground = inspection.Dialog != 0 ? inspection.Dialog : inspection.Windows
                .First(window => window.ProcessId == identity.ProcessId && window.ClassName == "XLMAIN").Handle;
            // Windows can refuse foreground activation even though revealing the window succeeded.
            _ = SetForegroundWindow(foreground);
        }
        return new(visible, visible && GetForegroundWindow() == foreground);
    }

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool ShowWindowAsync(nint window, int command);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsIconic(nint window);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(nint window, out uint processId);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetForegroundWindow(nint window);

    [DllImport("user32.dll")]
    private static extern nint GetForegroundWindow();
}
