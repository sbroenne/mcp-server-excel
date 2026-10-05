using System.ComponentModel;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.Core.Utilities;

/// <summary>
/// Uses physical window coordinates for an operation without changing the host's DPI settings.
/// </summary>
internal static class DpiAwareness
{
    private static readonly IntPtr PerMonitorAwareV2 = new(-4);

    public static T Execute<T>(Func<T> operation)
    {
        ArgumentNullException.ThrowIfNull(operation);

        IntPtr previous = SetThreadDpiAwarenessContext(PerMonitorAwareV2);
        if (previous == IntPtr.Zero)
        {
            throw new Win32Exception(
                Marshal.GetLastWin32Error(),
                "Could not enable per-monitor DPI awareness for Excel window coordinates.");
        }

        T result = default!;
        Exception? failure = null;
        try
        {
            result = operation();
        }
        catch (Exception exception)
        {
            failure = exception;
        }
        finally
        {
            if (SetThreadDpiAwarenessContext(previous) == IntPtr.Zero)
            {
                var restoreFailure = new Win32Exception(
                    Marshal.GetLastWin32Error(),
                    "Could not restore the thread DPI awareness context.");
                failure = failure is null
                    ? restoreFailure
                    : new AggregateException(
                        "The Excel window operation and DPI context restoration both failed.",
                        failure, restoreFailure);
            }
        }

        if (failure is not null)
        {
            ExceptionDispatchInfo.Capture(failure).Throw();
        }

        return result;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
}
