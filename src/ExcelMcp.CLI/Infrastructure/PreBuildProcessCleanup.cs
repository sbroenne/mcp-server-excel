using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

internal static class PreBuildProcessCleanup
{
    private static readonly TimeSpan GracefulRequestTimeout = TimeSpan.FromSeconds(3);
    private static readonly TimeSpan GracefulExitTimeout = TimeSpan.FromSeconds(12);

    internal static Task<OwnedProcessCleanup.CleanupResult> CleanupWithGracefulShutdownAsync(
        string pipeName,
        CancellationToken cancellationToken) =>
        CleanupWithGracefulShutdownAsync(
            pipeName,
            cancellationToken,
            token => RequestGracefulShutdownAsync(pipeName, token));

    internal static async Task<OwnedProcessCleanup.CleanupResult> CleanupWithGracefulShutdownAsync(
        string pipeName,
        CancellationToken cancellationToken,
        Func<CancellationToken, Task<bool>> requestGracefulShutdownAsync)
    {
        ArgumentNullException.ThrowIfNull(requestGracefulShutdownAsync);

        var snapshot = OwnedProcessCleanup.CaptureTrackedProcesses(pipeName);
        if (snapshot.TrackingStatus
            is DaemonProcessTracker.TrackingRecordStatus.Invalid
            or DaemonProcessTracker.TrackingRecordStatus.Unreadable)
        {
            return await OwnedProcessCleanup.CleanupAsync(
                pipeName,
                snapshot,
                cancellationToken);
        }

        _ = await requestGracefulShutdownAsync(cancellationToken);
        // The daemon may begin saving before its reply reaches the cleanup client.
        // Always allow that tracked generation to exit before force cleanup.
        _ = await WaitForTrackedGenerationExitAsync(
            pipeName,
            snapshot,
            cancellationToken);

        return await OwnedProcessCleanup.CleanupAsync(
            pipeName,
            snapshot,
            cancellationToken);
    }

    private static async Task<bool> RequestGracefulShutdownAsync(
        string pipeName,
        CancellationToken cancellationToken)
    {
        try
        {
            using var client = new ServiceClient(
                pipeName,
                connectTimeout: GracefulRequestTimeout,
                requestTimeout: GracefulRequestTimeout);
            var response = await client.SendAsync(
                new ServiceRequest { Command = "service.shutdown" },
                GracefulRequestTimeout,
                cancellationToken);
            return response.Success;
        }
        catch (IOException)
        {
            return false;
        }
        catch (TimeoutException)
        {
            return false;
        }
    }

    private static async Task<bool> WaitForTrackedGenerationExitAsync(
        string pipeName,
        OwnedProcessCleanup.ProcessSnapshot snapshot,
        CancellationToken cancellationToken)
    {
        if (!snapshot.DaemonMatched || snapshot.DaemonProcess is not { } expectedDaemon)
        {
            return true;
        }

        var deadline = DateTime.UtcNow + GracefulExitTimeout;
        while (DateTime.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var current = OwnedProcessCleanup.CaptureTrackedProcesses(pipeName);
            if (current.TrackingStatus
                is DaemonProcessTracker.TrackingRecordStatus.Invalid
                or DaemonProcessTracker.TrackingRecordStatus.Unreadable)
            {
                await Task.Delay(TimeSpan.FromMilliseconds(200), cancellationToken);
                continue;
            }

            var excelProcesses = snapshot.ExcelProcesses;
            if (current.DaemonProcess == expectedDaemon)
            {
                excelProcesses = excelProcesses
                    .Concat(current.ExcelProcesses)
                    .Distinct()
                    .ToList();
            }

            if (IsProcessExited(expectedDaemon)
                && excelProcesses.All(IsProcessExited))
            {
                return true;
            }

            await Task.Delay(TimeSpan.FromMilliseconds(200), cancellationToken);
        }

        return false;
    }

    private static bool IsProcessExited(
        DaemonProcessTracker.ProcessIdentity identity)
    {
        if (!DaemonProcessTracker.TryOpenMatchingProcess(identity, out var process))
        {
            return false;
        }

        process?.Dispose();
        return process == null;
    }
}
