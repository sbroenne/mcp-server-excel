using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>
/// Stops only the daemon and Excel processes recorded for one CLI pipe.
/// </summary>
internal static class OwnedProcessCleanup
{
    private static readonly TimeSpan ProcessExitTimeout = TimeSpan.FromSeconds(10);
    private const uint ProcessTerminate = 0x0001;
    private const uint Synchronize = 0x00100000;
    private const uint ProcessQueryLimitedInformation = 0x1000;
    private const uint WaitObject0 = 0x00000000;
    private const uint WaitTimeout = 0x00000102;
    private const int ErrorInvalidParameter = 87;

    internal sealed record ProcessSnapshot(
        DaemonProcessTracker.ProcessIdentity? DaemonProcess,
        bool DaemonMatched,
        IReadOnlyList<DaemonProcessTracker.ProcessIdentity> ExcelProcesses,
        DaemonProcessTracker.TrackingRecordStatus TrackingStatus,
        string? ErrorMessage);

    internal sealed record CleanupResult(
        bool Success,
        bool DaemonMatched,
        string? ErrorMessage);

    internal static ProcessSnapshot CaptureTrackedProcesses(string pipeName)
    {
        var readResult = DaemonProcessTracker.ReadProcessSnapshot(pipeName);
        if (readResult.Status != DaemonProcessTracker.TrackingRecordStatus.Available
            || readResult.Snapshot == null)
        {
            return new ProcessSnapshot(
                null,
                false,
                [],
                readResult.Status,
                readResult.ErrorMessage);
        }
        var trackedSnapshot = readResult.Snapshot;

        var daemonValidated = DaemonProcessTracker.TryOpenMatchingProcess(
            trackedSnapshot.DaemonProcess,
            out var daemonProcess);
        if (!daemonValidated)
        {
            return new ProcessSnapshot(
                trackedSnapshot.DaemonProcess,
                false,
                [],
                DaemonProcessTracker.TrackingRecordStatus.Unreadable,
                $"The tracked daemon process for CLI pipe '{pipeName}' could not be validated.");
        }
        var daemonMatched = daemonProcess != null;
        daemonProcess?.Dispose();

        var identities = new List<DaemonProcessTracker.ProcessIdentity>();
        foreach (var identity in trackedSnapshot.ExcelProcesses)
        {
            if (!DaemonProcessTracker.TryOpenMatchingProcess(identity, out var process))
            {
                return new ProcessSnapshot(
                    trackedSnapshot.DaemonProcess,
                    daemonMatched,
                    identities,
                    DaemonProcessTracker.TrackingRecordStatus.Unreadable,
                    $"Tracked Excel process {identity.ProcessId} for CLI pipe '{pipeName}' could not be validated.");
            }

            if (process != null)
            {
                process.Dispose();
                identities.Add(identity);
            }
        }

        return new ProcessSnapshot(
            trackedSnapshot.DaemonProcess,
            daemonMatched,
            identities,
            DaemonProcessTracker.TrackingRecordStatus.Available,
            null);
    }

    internal static async Task<CleanupResult> CleanupAsync(
        string pipeName,
        CancellationToken cancellationToken)
    {
        return await CleanupAsync(
            pipeName,
            CaptureTrackedProcesses(pipeName),
            cancellationToken);
    }

    internal static async Task<CleanupResult> CleanupAsync(
        string pipeName,
        ProcessSnapshot preShutdownSnapshot,
        CancellationToken cancellationToken)
    {
        if (preShutdownSnapshot.TrackingStatus
            is DaemonProcessTracker.TrackingRecordStatus.Invalid
            or DaemonProcessTracker.TrackingRecordStatus.Unreadable)
        {
            return new CleanupResult(
                false,
                false,
                preShutdownSnapshot.ErrorMessage
                ?? $"The daemon tracking record for CLI pipe '{pipeName}' is not usable.");
        }

        if (preShutdownSnapshot.TrackingStatus
            == DaemonProcessTracker.TrackingRecordStatus.Missing)
        {
            return new CleanupResult(true, false, null);
        }

        var success = true;

        if (preShutdownSnapshot.DaemonMatched
            && preShutdownSnapshot.DaemonProcess is { } daemonIdentity)
        {
            success &= await TryTerminateProcessAsync(daemonIdentity, cancellationToken);
        }

        IReadOnlyList<DaemonProcessTracker.ProcessIdentity> finalExcelProcesses = [];
        var finalRead = DaemonProcessTracker.ReadProcessSnapshot(pipeName);
        if (finalRead.Status
            is DaemonProcessTracker.TrackingRecordStatus.Invalid
            or DaemonProcessTracker.TrackingRecordStatus.Unreadable)
        {
            return new CleanupResult(
                false,
                preShutdownSnapshot.DaemonMatched,
                finalRead.ErrorMessage
                ?? $"The daemon tracking record for CLI pipe '{pipeName}' is not usable.");
        }

        if (preShutdownSnapshot.DaemonProcess is { } expectedDaemon
            && finalRead.Snapshot is { } finalSnapshot
            && finalSnapshot.DaemonProcess == expectedDaemon)
        {
            finalExcelProcesses = finalSnapshot.ExcelProcesses;
        }

        var excelProcesses = preShutdownSnapshot.ExcelProcesses
            .Concat(finalExcelProcesses)
            .Distinct()
            .ToList();
        foreach (var identity in excelProcesses)
        {
            success &= await TryTerminateProcessAsync(identity, cancellationToken);
        }

        if (success)
        {
            success = preShutdownSnapshot.DaemonProcess is not { } daemonToClear
                || DaemonProcessTracker.ClearIfDaemonMatches(pipeName, daemonToClear);
        }

        return new CleanupResult(
            success,
            preShutdownSnapshot.DaemonMatched,
            success
                ? null
                : $"One or more processes tracked for CLI pipe '{pipeName}' could not be stopped.");
    }

    internal static bool TryTerminateProcess(Process process, bool entireProcessTree)
    {
        try
        {
            if (!process.HasExited)
            {
                process.Kill(entireProcessTree);
            }

            return true;
        }
        catch (InvalidOperationException)
        {
            return true;
        }
        catch (Win32Exception)
        {
            try
            {
                process.Refresh();
                return process.HasExited;
            }
            catch (InvalidOperationException)
            {
                return true;
            }
            catch (Win32Exception)
            {
                return false;
            }
        }
        catch (NotSupportedException)
        {
            return false;
        }
    }

    private static async Task<bool> TryTerminateProcessAsync(
        DaemonProcessTracker.ProcessIdentity identity,
        CancellationToken cancellationToken)
    {
        if (!TryOpenMatchingProcessHandle(identity, out var processHandle))
        {
            return false;
        }

        if (processHandle == null)
        {
            return true;
        }

        using (processHandle)
        {
            try
            {
                if (!TerminateProcess(processHandle, 1)
                    && WaitForSingleObject(processHandle, 0) != WaitObject0)
                {
                    return false;
                }

                var deadline = DateTime.UtcNow + ProcessExitTimeout;
                while (DateTime.UtcNow < deadline)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var waitResult = WaitForSingleObject(processHandle, 100);
                    if (waitResult == WaitObject0)
                    {
                        return true;
                    }

                    if (waitResult != WaitTimeout)
                    {
                        return false;
                    }

                    await Task.Yield();
                }

                return false;
            }
            catch (OperationCanceledException)
            {
                return false;
            }
        }
    }

    private static bool TryOpenMatchingProcessHandle(
        DaemonProcessTracker.ProcessIdentity identity,
        out SafeProcessHandle? processHandle)
    {
        processHandle = OpenProcess(
            ProcessTerminate | Synchronize | ProcessQueryLimitedInformation,
            inheritHandle: false,
            identity.ProcessId);
        if (processHandle.IsInvalid)
        {
            var error = Marshal.GetLastWin32Error();
            processHandle.Dispose();
            processHandle = null;
            return error == ErrorInvalidParameter;
        }

        if (!GetProcessTimes(
                processHandle,
                out var creationTime,
                out _,
                out _,
                out _))
        {
            processHandle.Dispose();
            processHandle = null;
            return false;
        }

        if (creationTime != identity.StartedAtUtcFileTime
            || WaitForSingleObject(processHandle, 0) == WaitObject0)
        {
            processHandle.Dispose();
            processHandle = null;
        }

        return true;
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
    private static extern bool TerminateProcess(SafeProcessHandle processHandle, uint exitCode);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint WaitForSingleObject(SafeProcessHandle handle, uint milliseconds);
}
