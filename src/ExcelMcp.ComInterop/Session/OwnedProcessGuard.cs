using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Validates and operates on one owned process through the same native handle.
/// </summary>
internal static class OwnedProcessGuard
{
    private const uint ProcessTerminate = 0x0001;
    private const uint Synchronize = 0x00100000;
    private const uint ProcessQueryLimitedInformation = 0x1000;
    private const uint WaitObject0 = 0x00000000;
    private const uint WaitTimeout = 0x00000102;
    private const int ErrorInvalidParameter = 87;

    internal static bool IsAlive(ExcelProcessIdentity identity)
    {
        var probe = ProbeMatchingProcess(
            identity,
            ProcessQueryLimitedInformation | Synchronize,
            out var handle);
        using (handle)
        {
            return IsAlive(probe);
        }
    }

    internal static bool IsAlive(ProcessIdentityProbe probe) =>
        probe != ProcessIdentityProbe.ConfirmedExited;

    internal static bool TryConfirmExited(ExcelProcessIdentity identity)
    {
        var probe = ProbeMatchingProcess(
            identity,
            ProcessQueryLimitedInformation | Synchronize,
            out var handle);
        using (handle)
        {
            return probe == ProcessIdentityProbe.ConfirmedExited;
        }
    }

    internal static bool TryTerminate(
        ExcelProcessIdentity identity,
        TimeSpan waitBeforeTermination,
        TimeSpan waitAfterTermination,
        out bool terminated)
    {
        terminated = false;
        var probe = ProbeMatchingProcess(
            identity,
            ProcessTerminate | ProcessQueryLimitedInformation | Synchronize,
            out var handle);
        if (probe == ProcessIdentityProbe.Indeterminate)
        {
            return false;
        }

        if (probe == ProcessIdentityProbe.ConfirmedExited)
        {
            return true;
        }

        if (handle == null)
        {
            return false;
        }

        using (handle)
        {
            var initialWait = WaitForSingleObject(
                handle,
                ToWaitMilliseconds(waitBeforeTermination));
            if (initialWait == WaitObject0)
            {
                return true;
            }

            if (initialWait != WaitTimeout)
            {
                return false;
            }

            if (!TerminateProcess(handle, 1)
                && WaitForSingleObject(handle, 0) != WaitObject0)
            {
                return false;
            }

            terminated = true;
            return WaitForSingleObject(
                handle,
                ToWaitMilliseconds(waitAfterTermination)) == WaitObject0;
        }
    }

    private static ProcessIdentityProbe ProbeMatchingProcess(
        ExcelProcessIdentity identity,
        uint desiredAccess,
        out SafeProcessHandle? handle)
    {
        handle = OpenProcess(desiredAccess, inheritHandle: false, identity.ProcessId);
        if (handle.IsInvalid)
        {
            var error = Marshal.GetLastWin32Error();
            handle.Dispose();
            handle = null;
            return error == ErrorInvalidParameter
                ? ProcessIdentityProbe.ConfirmedExited
                : ProcessIdentityProbe.Indeterminate;
        }

        if (!GetProcessTimes(handle, out var creationTime, out _, out _, out _))
        {
            handle.Dispose();
            handle = null;
            return ProcessIdentityProbe.Indeterminate;
        }

        if (creationTime != identity.StartedAtUtcFileTime
            || WaitForSingleObject(handle, 0) == WaitObject0)
        {
            handle.Dispose();
            handle = null;
            return ProcessIdentityProbe.ConfirmedExited;
        }

        return ProcessIdentityProbe.Alive;
    }

    internal enum ProcessIdentityProbe
    {
        Alive,
        ConfirmedExited,
        Indeterminate
    }

    private static uint ToWaitMilliseconds(TimeSpan timeout)
    {
        if (timeout <= TimeSpan.Zero)
        {
            return 0;
        }

        return (uint)Math.Min(timeout.TotalMilliseconds, uint.MaxValue - 1);
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern SafeProcessHandle OpenProcess(
        uint desiredAccess,
        [MarshalAs(UnmanagedType.Bool)] bool inheritHandle,
        int processId);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetProcessTimes(
        SafeProcessHandle processHandle,
        out long creationTime,
        out long exitTime,
        out long kernelTime,
        out long userTime);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool TerminateProcess(
        SafeProcessHandle processHandle,
        uint exitCode);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint WaitForSingleObject(
        SafeProcessHandle handle,
        uint milliseconds);
}
