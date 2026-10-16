using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
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

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Win32 readiness probe handles only process/window inspection failures, logs the original exception, and returns Unknown so save and close remain blocked rather than reporting readiness.")]
    internal static WorkbookRefreshState Read(ExcelProcessIdentity? identity, ILogger logger)
    {
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
                    var titleBuffer = new char[512];
                    var titleLength = GetWindowText(window, titleBuffer, titleBuffer.Length);
                    if (titleLength > 0)
                    {
                        title = new string(titleBuffer, 0, titleLength);
                    }
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
            var state = Classify(ownedProcess.ProcessId, windows);
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
    {
        var roots = windows.Where(window => window.ProcessId == processId && window.ClassName == "XLMAIN").ToArray();
        if (roots.Length == 0) return WorkbookRefreshState.Unknown;
        var byHandle = windows.ToDictionary(window => window.Handle);
        foreach (var root in roots)
        {
            if (root.Enabled) continue;
            foreach (var window in windows)
            {
                if (!window.Visible || !window.Enabled || window.Handle == root.Handle) continue;
                var owner = window.Owner;
                for (var depth = 0; owner != nint.Zero && depth < windows.Count; depth++)
                {
                    if (owner == root.Handle)
                    {
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

    private static bool IsNativeQueryApprovalTitle(string title) =>
        NativeQueryApprovalTitleFragments.Any(fragment =>
            title.Contains(fragment, StringComparison.OrdinalIgnoreCase));

    private delegate bool WindowCallback(nint window, nint data);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumWindows(WindowCallback callback, nint data);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(nint window, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(nint window, [Out] char[] className, int capacity);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(nint window, [Out] char[] text, int capacity);

    [DllImport("user32.dll")]
    private static extern nint GetWindow(nint window, uint command);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowVisible(nint window);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowEnabled(nint window);
}
