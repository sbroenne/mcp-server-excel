using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal static class ExcelDialogProbe
{
    internal readonly record struct WindowSnapshot(
        nint Handle, nint Owner, int ProcessId, bool Visible, bool Enabled, string ClassName, string Title = "");

    internal static WorkbookRefreshState Read(ExcelProcessIdentity? identity, ILogger logger) =>
        Inspect(identity, logger).State;

    internal static string Reveal(ExcelProcessIdentity? identity, ILogger logger)
    {
        if (identity is not { } owned || !OwnedProcessGuard.MatchesLiveProcess(owned))
        {
            logger.LogWarning("Cannot reveal Excel without a live owned process identity");
            return "unavailable";
        }

        var outcome = "unavailable";
        var enumerated = EnumWindows((window, _) =>
        {
            if (GetWindowThreadProcessId(window, out var processId) == 0 ||
                processId != owned.ProcessId)
                return true;
            var buffer = new char[256];
            var length = GetClassName(window, buffer, buffer.Length);
            if (new string(buffer, 0, length) != "XLMAIN")
                return true;
            if (IsWindowVisible(window))
            {
                outcome = "alreadyVisible";
                return false;
            }
            if (!OwnedProcessGuard.MatchesLiveProcess(owned))
                return false;
            if (ShowWindowAsync(window, 5))
                outcome = IsWindowVisible(window) ? "shown" : "showRequested";
            return false;
        }, nint.Zero);
        if (!enumerated && outcome == "unavailable")
            logger.LogWarning("Excel window could not be revealed; workbook ownership is retained");
        return outcome;
    }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool ShowWindowAsync(nint window, int command);

    internal readonly record struct Inspection(
        WorkbookRefreshState State, IReadOnlyList<WindowSnapshot> Windows, nint Dialog);

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Win32 readiness probe handles only process/window inspection failures, logs the original exception, and returns Unknown so save and close remain blocked rather than reporting readiness.")]
    internal static Inspection Inspect(ExcelProcessIdentity? identity, ILogger logger)
    {
        if (identity is not { } ownedProcess)
        {
            logger.LogWarning("Cannot inspect Excel dialogs without an owned process identity");
            return new(WorkbookRefreshState.Unknown, [], 0);
        }

        try
        {
            using var process = Process.GetProcessById(ownedProcess.ProcessId);
            if (process.HasExited || process.StartTime.ToUniversalTime().ToFileTimeUtc() != ownedProcess.StartedAtUtcFileTime)
            {
                logger.LogWarning("Excel dialog inspection could not confirm the owned process identity");
                return new(WorkbookRefreshState.Unknown, [], 0);
            }

            var windows = new List<WindowSnapshot>();
            var classReadFailed = false;
            var success = EnumWindows((window, _) =>
            {
                if (GetWindowThreadProcessId(window, out var processId) == 0) return true;
                var className = string.Empty;
                if (processId == ownedProcess.ProcessId)
                {
                    var buffer = new char[256];
                    var length = GetClassName(window, buffer, buffer.Length);
                    if (length == 0)
                    {
                        classReadFailed = true;
                        return true;
                    }
                    className = new string(buffer, 0, length);
                }
                windows.Add(new WindowSnapshot(window, GetWindow(window, 4),
                    checked((int)processId), IsWindowVisible(window), IsWindowEnabled(window), className));
                return true;
            }, nint.Zero);

            if (!success || classReadFailed || process.HasExited)
            {
                logger.LogWarning("Excel dialog window inspection failed; readiness remains unknown");
                return new(WorkbookRefreshState.Unknown, [], 0);
            }
            var state = Classify(ownedProcess.ProcessId, windows, out var dialog);
            if (state == WorkbookRefreshState.Unknown)
            {
                logger.LogWarning("Excel's main window could not be found; dialog readiness remains unknown");
            }
            return new(state, windows, dialog);
        }
        catch (Exception ex) when (ex is Win32Exception or ArgumentException or InvalidOperationException)
        {
            logger.LogWarning(ex, "Could not inspect Excel dialogs; readiness remains unknown");
            return new(WorkbookRefreshState.Unknown, [], 0);
        }
    }

    internal static WorkbookRefreshState Classify(int processId, IReadOnlyList<WindowSnapshot> windows)
        => Classify(processId, windows, out _);

    internal static WorkbookRefreshState Classify(
        int processId, IReadOnlyList<WindowSnapshot> windows, out nint dialog)
    {
        dialog = nint.Zero;
        var roots = windows.Where(window => window.ProcessId == processId && window.ClassName == "XLMAIN").ToArray();
        if (roots.Length == 0) return WorkbookRefreshState.Unknown;
        var byHandle = windows.ToDictionary(window => window.Handle);
        foreach (var root in roots)
        {
            foreach (var window in windows)
            {
                if (!window.Visible || !window.Enabled || window.Handle == root.Handle) continue;
                if (root.Enabled) continue;
                var owner = window.Owner;
                for (var depth = 0; owner != nint.Zero && depth < windows.Count; depth++)
                {
                    if (owner == root.Handle)
                    {
                        dialog = window.Handle;
                        return WorkbookRefreshState.DialogOpen;
                    }
                    if (!byHandle.TryGetValue(owner, out var parent)) break;
                    owner = parent.Owner;
                }
            }
        }
        return WorkbookRefreshState.Ready;
    }

    private delegate bool WindowCallback(nint window, nint data);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumWindows(WindowCallback callback, nint data);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(nint window, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(nint window, [Out] char[] className, int capacity);

    [DllImport("user32.dll")]
    private static extern nint GetWindow(nint window, uint command);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowVisible(nint window);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowEnabled(nint window);
}
