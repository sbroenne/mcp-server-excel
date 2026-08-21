using System.Text.Json;
using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

internal static class Program
{
    private static readonly TimeSpan GracefulRequestTimeout = TimeSpan.FromSeconds(3);
    private static readonly TimeSpan GracefulExitTimeout = TimeSpan.FromSeconds(12);

    private static async Task<int> Main()
    {
        var pipeName = Environment.GetEnvironmentVariable("EXCELMCP_CLI_PIPE")
            ?? ServiceSecurity.GetCliPipeName();
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(45));

        try
        {
            var result = await DaemonStartupLock.WithLockAsync(
                pipeName,
                () => CleanupWithGracefulShutdownAsync(pipeName, timeout.Token),
                timeout.Token);
            Console.WriteLine(JsonSerializer.Serialize(new
            {
                success = result.Success,
                daemonMatched = result.DaemonMatched,
                error = result.ErrorMessage
            }));
            return result.Success ? 0 : 1;
        }
        catch (Exception ex)
        {
            Console.WriteLine(JsonSerializer.Serialize(new
            {
                success = false,
                error = ex.Message
            }));
            return 1;
        }
    }

    private static async Task<OwnedProcessCleanup.CleanupResult> CleanupWithGracefulShutdownAsync(
        string pipeName,
        CancellationToken cancellationToken)
    {
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

        var gracefulAcknowledged = false;
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
            gracefulAcknowledged = response.Success;
        }
        catch (IOException)
        {
            gracefulAcknowledged = false;
        }
        catch (TimeoutException)
        {
            gracefulAcknowledged = false;
        }

        if (gracefulAcknowledged)
        {
            _ = await WaitForTrackedDaemonExitAsync(
                pipeName,
                snapshot,
                cancellationToken);
        }

        return await OwnedProcessCleanup.CleanupAsync(
            pipeName,
            snapshot,
            cancellationToken);
    }

    private static async Task<bool> WaitForTrackedDaemonExitAsync(
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
            if (current.TrackingStatus == DaemonProcessTracker.TrackingRecordStatus.Missing
                || !current.DaemonMatched
                || current.DaemonProcess != expectedDaemon)
            {
                return true;
            }

            await Task.Delay(TimeSpan.FromMilliseconds(200), cancellationToken);
        }

        return false;
    }
}
