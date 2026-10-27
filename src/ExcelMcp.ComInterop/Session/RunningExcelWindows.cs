using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Principal;
using Microsoft.Win32.SafeHandles;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal static class RunningExcelWindows
{
    internal sealed record ApplicationWindow(nint Handle, ExcelProcessIdentity Process, IReadOnlyList<nint>? AlternateHandles = null);
    internal sealed record Discovery(IReadOnlyList<ApplicationWindow> Windows, IReadOnlyList<WorkbookDiscoveryError> Errors);
    internal static Func<Discovery>? FindHookForTests { get; set; }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Exceptions are contained at the native callback boundary and returned as explicit discovery errors alongside any successfully located Excel applications.")]
    internal static Discovery Find()
    {
        if (FindHookForTests is { } hook)
            return hook();

        var windows = new Dictionary<int, ApplicationWindow>();
        var errors = new List<WorkbookDiscoveryError>();
        int? currentProcessId = null;

        try
        {
            using var current = Process.GetCurrentProcess();
            using var identity = WindowsIdentity.GetCurrent();
            Marshal.SetLastPInvokeError(0);
            var completed = EnumWindows((window, parameter) =>
            {
                int? windowProcessId = null;
                try
                {
                    if (ReadClass(window) != "XLMAIN")
                        return true;

                    _ = GetWindowThreadProcessId(window, out var processId);
                    if (processId == 0 || processId > int.MaxValue)
                    {
                        errors.Add(new WorkbookDiscoveryError(
                            null, "An Excel window disappeared before its process could be identified."));
                        return true;
                    }

                    windowProcessId = (int)processId;
                    currentProcessId = windowProcessId;
                    using var process = Process.GetProcessById(windowProcessId.Value);
                    if (process.SessionId != current.SessionId)
                        return true;

                    if (!OpenProcessToken(process.SafeHandle, 0x0008, out var token))
                    {
                        errors.Add(new WorkbookDiscoveryError(
                            process.Id, new Win32Exception(Marshal.GetLastWin32Error()).Message));
                        return true;
                    }

                    using (token)
                    using (var owner = new WindowsIdentity(token.DangerousGetHandle()))
                    {
                        if (owner.User != identity.User)
                            return true;
                    }

                    nint workbookWindow = nint.Zero;
                    Exception? childEnumerationFailure = null;
                    var childEnumerationCompleted = EnumChildWindows(window, (child, parameter) =>
                    {
                        try
                        {
                            if (workbookWindow != nint.Zero)
                                return true;
                            if (ReadClass(child) == "EXCEL7")
                            {
                                workbookWindow = child;
                            }
                            return true;
                        }
                        catch (Exception ex)
                        {
                            childEnumerationFailure = ex;
                            return false;
                        }
                    }, nint.Zero);

                    if (childEnumerationFailure is not null)
                    {
                        errors.Add(new WorkbookDiscoveryError(process.Id, Describe(childEnumerationFailure)));
                        return true;
                    }
                    if (!childEnumerationCompleted)
                    {
                        var errorCode = Marshal.GetLastWin32Error();
                        errors.Add(new WorkbookDiscoveryError(
                            process.Id,
                            errorCode == 0
                                ? "Windows could not enumerate the child windows of an Excel window."
                                : new Win32Exception(errorCode).Message));
                        if (workbookWindow == nint.Zero)
                            return true;
                    }

                    if (workbookWindow != nint.Zero)
                    {
                        if (windows.TryGetValue(process.Id, out var existing))
                            windows[process.Id] = existing with
                            {
                                AlternateHandles = [.. existing.AlternateHandles ?? [], workbookWindow]
                            };
                        else
                            windows.Add(process.Id, new ApplicationWindow(
                                workbookWindow,
                                new ExcelProcessIdentity(
                                    process.Id, process.StartTime.ToUniversalTime().ToFileTimeUtc())));
                    }
                    return true;
                }
                catch (Exception ex)
                {
                    errors.Add(new WorkbookDiscoveryError(windowProcessId, Describe(ex)));
                    return true;
                }

            }, nint.Zero);

            if (!completed)
            {
                var errorCode = Marshal.GetLastPInvokeError();
                errors.Add(new WorkbookDiscoveryError(
                    currentProcessId,
                    errorCode == 0
                        ? "Windows stopped enumerating Excel windows before discovery completed."
                        : new Win32Exception(errorCode).Message));
            }
        }
        catch (Exception ex)
        {
            errors.Add(new WorkbookDiscoveryError(null, Describe(ex)));
        }

        return new Discovery(windows.Values.ToArray(), errors);
    }

    internal static Excel.Application Connect(ApplicationWindow window)
    {
        if (!OwnedProcessGuard.MatchesLiveProcess(window.Process))
            throw new InvalidOperationException("The selected Excel process is no longer running.");
        var dispatchId = new Guid("00020400-0000-0000-C000-000000000046");
        foreach (var handle in new[] { window.Handle }.Concat(window.AlternateHandles ?? []))
        {
            _ = GetWindowThreadProcessId(handle, out var processId);
            if (processId != window.Process.ProcessId)
                throw new InvalidOperationException("The selected Excel window no longer belongs to the expected process.");
            object? accessible = null;
            Excel.Application? application = null;
            try
            {
                Marshal.ThrowExceptionForHR(AccessibleObjectFromWindow(
                    handle, 0xFFFFFFF0, ref dispatchId, out accessible));
                if (accessible is Excel.ProtectedViewWindow)
                    continue;
                application = accessible is Excel.Window workbookWindow
                    ? workbookWindow.Application
                    : throw new InvalidOperationException("The selected Excel window exposes no supported workbook interface.");
                var result = application;
                application = null;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref application);
                ComUtilities.Release(ref accessible);
            }
        }
        throw new InvalidOperationException(
            "Only a restricted Protected View window could be reached. Respond in Excel, or use an already connected normal Excel window; no trust was changed.");
    }

    private static string Describe(Exception exception) =>
        $"{exception.GetType().Name}: {exception.Message}";

    private static string ReadClass(nint window)
    {
        var buffer = new char[256];
        var length = GetClassName(window, buffer, buffer.Length);
        if (length > 0)
            return new string(buffer, 0, length);

        var errorCode = Marshal.GetLastWin32Error();
        if (errorCode == 1400) // The window disappeared during enumeration.
            return string.Empty;
        throw errorCode == 0
            ? new InvalidOperationException("Windows returned no class name for an enumerated window.")
            : new Win32Exception(errorCode);
    }

    private delegate bool WindowCallback(nint window, nint parameter);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumWindows(WindowCallback callback, nint parameter);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumChildWindows(nint parent, WindowCallback callback, nint parameter);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(nint window, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int GetClassName(nint window, [Out] char[] className, int capacity);

    [DllImport("oleacc.dll")]
    private static extern int AccessibleObjectFromWindow(
        nint window, uint objectId, ref Guid interfaceId,
        [MarshalAs(UnmanagedType.Interface)] out object? result);

    [DllImport("advapi32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool OpenProcessToken(
        SafeProcessHandle process, uint access, out SafeAccessTokenHandle token);
}
