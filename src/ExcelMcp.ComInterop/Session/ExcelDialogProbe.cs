using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Extensions.Logging;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal static class ExcelDialogProbe
{
    internal readonly record struct WindowSnapshot(
        nint Handle, nint Owner, int ProcessId, bool Visible, bool Enabled, string ClassName, string Title = "");

    internal static WorkbookRefreshState Read(
        ExcelProcessIdentity? identity, ILogger logger, bool includeEnabledOwnerOfficeDialog = false) =>
        Inspect(identity, logger, includeEnabledOwnerOfficeDialog).State;

    internal static string ReadDetails(
        ExcelProcessIdentity? identity, ILogger logger, bool includeEnabledOwnerOfficeDialog = false)
    {
        var inspection = Inspect(identity, logger, includeEnabledOwnerOfficeDialog);
        if (inspection.State == WorkbookRefreshState.Ready)
            return "No blocking Excel dialog was detected.";
        if (inspection.State == WorkbookRefreshState.Unknown)
            return "Excel dialog inspection could not confirm the session's window ownership.";
        if (GetWindowThreadProcessId(inspection.Dialog, out var dialogProcessId) == 0)
            return "Dialog text could not be read because the owned blocking window is no longer confirmed.";

        var text = new StringBuilder();
        var started = Stopwatch.GetTimestamp();
        var incomplete = false;
        var labels = 0;
        var controls = 0;

        bool ReadWindow(nint window)
        {
            var remainingMilliseconds = 500 - (int)Stopwatch.GetElapsedTime(started).TotalMilliseconds;
            if (remainingMilliseconds <= 0 || text.Length >= 8192)
            {
                incomplete = true;
                return false;
            }
            if (GetWindowThreadProcessId(window, out var processId) == 0 || processId != dialogProcessId ||
                (window != inspection.Dialog && !IsChild(inspection.Dialog, window)))
            {
                incomplete = true;
                return true;
            }
            var buffer = new char[2048];
            if (SendMessageTimeout(window, 0x000D, (nuint)buffer.Length, buffer,
                    0x0023, (uint)Math.Min(50, remainingMilliseconds), out var length) == nint.Zero)
            {
                incomplete = true;
                return true;
            }
            var value = new string(buffer, 0, (int)Math.Min(length, (nuint)(buffer.Length - 1))).Trim();
            if (value.Length == 0) return true;
            if (window != inspection.Dialog) labels++;
            if (text.Length > 0 && text.Length < 8192) text.Append('\n');
            var remaining = 8192 - text.Length;
            text.Append(value.AsSpan(0, Math.Min(value.Length, remaining)));
            incomplete |= value.Length > remaining || length >= (nuint)(buffer.Length - 1);
            return true;
        }

        ReadWindow(inspection.Dialog);
        EnumChildWindows(inspection.Dialog, (window, _) =>
        {
            if (++controls > 64 || Stopwatch.GetElapsedTime(started) >= TimeSpan.FromMilliseconds(500))
            {
                incomplete = true;
                return false;
            }
            if (!IsWindowVisible(window)) return true;
            var className = new char[256];
            var length = GetClassName(window, className, className.Length);
            if (length == 0)
            {
                incomplete = true;
                return true;
            }
            var name = new string(className, 0, length);
            if (!name.Equals("Static", StringComparison.OrdinalIgnoreCase) &&
                !name.Equals("Button", StringComparison.OrdinalIgnoreCase))
                return true;
            return ReadWindow(window);
        }, nint.Zero);

        var confirmed = Inspect(identity, logger, includeEnabledOwnerOfficeDialog);
        if (confirmed.State != WorkbookRefreshState.DialogOpen || confirmed.Dialog != inspection.Dialog)
            return "Dialog text could not be read because the owned blocking window changed during inspection.";
        if (text.Length == 0)
            return "Dialog text could not be read: Windows exposed no readable title or message labels within the bounded read.";
        return "Dialog text: " + text +
            (incomplete ? "\n[Some dialog text could not be read or was truncated.]" :
                labels == 0 ? "\n[No readable message labels were exposed.]" : "");
    }

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
    internal static Inspection Inspect(
        ExcelProcessIdentity? identity, ILogger logger, bool includeEnabledOwnerOfficeDialog = false)
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
            var state = Classify(ownedProcess.ProcessId, windows, out var dialog, includeEnabledOwnerOfficeDialog);
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
        int processId,
        IReadOnlyList<WindowSnapshot> windows,
        out nint dialog,
        bool includeEnabledOwnerOfficeDialog = false)
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
                if (root.Enabled &&
                    (!includeEnabledOwnerOfficeDialog || window.ProcessId != processId ||
                     window.ClassName != "bosa_sdm_XL9"))
                    continue;
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
    private static extern bool EnumChildWindows(nint parent, WindowCallback callback, nint data);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsChild(nint parent, nint window);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern nint SendMessageTimeout(nint window, uint message, nuint wParam,
        [Out] char[] text, uint flags, uint timeoutMilliseconds, out nuint result);

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
