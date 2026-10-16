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

    private static readonly string[] NativeQueryApprovalTitleFragments =
    [
        "Native Database Query",
        "Native Datenbankabfrage",
        "Requête de base de données native",
        "Consulta de base de datos nativa",
        "Consulta de base de dados nativa",
        "Consulta de banco de dados nativa",
        "Query di database nativa",
        "ネイティブ データベース クエリ",
        "네이티브 데이터베이스 쿼리",
        "本机数据库查询",
        "原生資料庫查詢"
    ];

    internal static WorkbookRefreshState Read(ExcelProcessIdentity? identity, ILogger logger) =>
        Read(identity, logger, includeText: false, out _);

    internal static WorkbookRefreshState ReadWithText(
        ExcelProcessIdentity? identity, ILogger logger, out string? dialogText) =>
        Read(identity, logger, includeText: true, out dialogText);

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Win32 readiness probe handles only process/window inspection failures, logs the original exception, and returns Unknown so save and close remain blocked rather than reporting readiness.")]
    private static WorkbookRefreshState Read(
        ExcelProcessIdentity? identity, ILogger logger, bool includeText, out string? dialogText)
    {
        dialogText = null;
        if (identity is not { } ownedProcess)
        {
            logger.LogWarning("Cannot inspect Excel dialogs without an owned process identity");
            return WorkbookRefreshState.Unknown;
        }

        try
        {
            using var process = Process.GetProcessById(ownedProcess.ProcessId);
            if (process.HasExited || process.StartTime.ToUniversalTime().ToFileTimeUtc() != ownedProcess.StartedAtUtcFileTime)
            {
                logger.LogWarning("Excel dialog inspection could not confirm the owned process identity");
                return WorkbookRefreshState.Unknown;
            }

            var windows = new List<WindowSnapshot>();
            var classReadFailed = false;
            var success = EnumWindows((window, _) =>
            {
                if (GetWindowThreadProcessId(window, out var processId) == 0) return true;
                var className = string.Empty;
                var title = string.Empty;
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
                    title = ReadWindowTitle(window);
                }
                windows.Add(new WindowSnapshot(window, GetWindow(window, 4),
                    checked((int)processId), IsWindowVisible(window), IsWindowEnabled(window), className, title));
                return true;
            }, nint.Zero);

            if (!success || classReadFailed || process.HasExited)
            {
                logger.LogWarning("Excel dialog window inspection failed; readiness remains unknown");
                return WorkbookRefreshState.Unknown;
            }
            var state = Classify(ownedProcess.ProcessId, windows, includeEnabledOwnerOfficeDialog: includeText, out var dialog);
            if (includeText && dialog != nint.Zero)
            {
                dialogText = ReadDialogText(dialog);
            }
            if (state == WorkbookRefreshState.Unknown)
            {
                logger.LogWarning("Excel's main window could not be found; dialog readiness remains unknown");
            }
            return state;
        }
        catch (Exception ex) when (ex is Win32Exception or ArgumentException or InvalidOperationException)
        {
            logger.LogWarning(ex, "Could not inspect Excel dialogs; readiness remains unknown");
            return WorkbookRefreshState.Unknown;
        }
    }

    internal static WorkbookRefreshState Classify(int processId, IReadOnlyList<WindowSnapshot> windows)
        => Classify(processId, windows, includeEnabledOwnerOfficeDialog: false, out _);

    internal static WorkbookRefreshState Classify(
        int processId, IReadOnlyList<WindowSnapshot> windows, bool includeEnabledOwnerOfficeDialog, out nint dialog)
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
                // Excel's InputBox uses an Office dialog without disabling its XLMAIN owner.
                if (root.Enabled && (!includeEnabledOwnerOfficeDialog ||
                    window.ProcessId != processId || window.ClassName != "bosa_sdm_XL9"))
                    continue;
                var owner = window.Owner;
                for (var depth = 0; owner != nint.Zero && depth < windows.Count; depth++)
                {
                    if (owner == root.Handle)
                    {
                        dialog = window.Handle;
                        return IsNativeQueryApprovalTitle(window.Title)
                            ? WorkbookRefreshState.NativeQueryApproval
                            : WorkbookRefreshState.DialogOpen;
                    }
                    if (!byHandle.TryGetValue(owner, out var parent)) break;
                    owner = parent.Owner;
                }
            }
        }
        return WorkbookRefreshState.Ready;
    }

    private static string ReadDialogText(nint dialog)
    {
        var text = new StringBuilder();
        var started = Stopwatch.GetTimestamp();
        var incomplete = false;
        var controls = 0;

        bool ReadWindow(nint window)
        {
            if (Stopwatch.GetElapsedTime(started) >= TimeSpan.FromSeconds(1) ||
                text.Length >= 8192 || ++controls > 64)
            {
                incomplete = true;
                return false;
            }
            var buffer = new char[2048];
            // Unlike an ordinary SendMessage, this cannot wait indefinitely for a hung dialog.
            if (SendMessageTimeout(window, 0x000D, (nuint)buffer.Length, buffer,
                    0x0003, 100, out var length) == nint.Zero)
            {
                incomplete = true;
                return true;
            }
            var value = new string(buffer, 0, (int)Math.Min(length, (nuint)(buffer.Length - 1))).Trim();
            if (value.Length > 0)
            {
                if (text.Length > 0)
                {
                    if (text.Length + Environment.NewLine.Length >= 8192)
                    {
                        incomplete = true;
                        return false;
                    }
                    text.AppendLine();
                }
                var remaining = 8192 - text.Length;
                text.Append(value.AsSpan(0, Math.Min(value.Length, remaining)));
                incomplete |= value.Length > remaining || length >= (nuint)(buffer.Length - 1);
            }
            return true;
        }

        ReadWindow(dialog);
        var titleLength = text.Length;
        EnumChildWindows(dialog, (window, _) =>
        {
            if (!IsWindowVisible(window)) return true;
            var classBuffer = new char[256];
            var classLength = GetClassName(window, classBuffer, classBuffer.Length);
            if (classLength == 0)
            {
                incomplete = true;
                return true;
            }
            // Read labels and messages, not editable values such as sign-in credentials.
            var className = new string(classBuffer, 0, classLength);
            return (!className.Equals("Static", StringComparison.OrdinalIgnoreCase) &&
                    !className.Equals("Button", StringComparison.OrdinalIgnoreCase)) || ReadWindow(window);
        }, nint.Zero);
        if (incomplete) text.AppendLine().Append("[Some dialog text could not be read or was truncated.]");
        else if (text.Length == titleLength)
            text.AppendLine().Append("[The dialog did not expose readable message text.]");
        return text.ToString();
    }

    private static bool IsNativeQueryApprovalTitle(string title) =>
        NativeQueryApprovalTitleFragments.Any(fragment =>
            title.Contains(fragment, StringComparison.OrdinalIgnoreCase));

    private static string ReadWindowTitle(nint window)
    {
        var buffer = new char[512];
        if (SendMessageTimeout(window, 0x000D, (nuint)buffer.Length, buffer,
                0x0003, 100, out var length) == nint.Zero)
            return string.Empty;
        return new string(buffer, 0, (int)Math.Min(length, (nuint)(buffer.Length - 1)));
    }

    private delegate bool WindowCallback(nint window, nint data);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumChildWindows(nint parent, WindowCallback callback, nint data);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern nint SendMessageTimeout(nint window, uint message, nuint wParam,
        [Out] char[] lParam, uint flags, uint timeoutMilliseconds, out nuint result);

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
