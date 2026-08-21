using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Service;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

// ============================================================================
// SERVICE LIFECYCLE COMMANDS
// ============================================================================

/// <summary>
/// Starts the ExcelMCP CLI Service daemon if not already running.
/// Launches a background process running "excelcli service run".
/// </summary>
internal sealed class ServiceStartCommand : AsyncCommand
{
    protected override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
    {
        try
        {
            using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = "Service started." }, ServiceProtocol.JsonOptions));
            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = ex.Message }, ServiceProtocol.JsonOptions));
            return 1;
        }
    }
}

/// <summary>
/// Stops the ExcelMCP CLI Service daemon and Excel processes tracked for its selected pipe.
/// </summary>
internal sealed class ServiceStopCommand : AsyncCommand
{
    private static readonly TimeSpan CommandTimeout = TimeSpan.FromSeconds(2);
    private static readonly TimeSpan ShutdownWaitTimeout = TimeSpan.FromSeconds(10);
    private static readonly TimeSpan ShutdownPollInterval = TimeSpan.FromMilliseconds(250);

    protected override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
    {
        var pipeName = DaemonAutoStart.GetPipeName();
        return await DaemonAutoStart.WithStartupLockAsync(
            pipeName,
            () => ExecuteWithStartupLockAsync(pipeName, cancellationToken),
            cancellationToken);
    }

    private static async Task<int> ExecuteWithStartupLockAsync(
        string pipeName,
        CancellationToken cancellationToken)
    {
        var preShutdownSnapshot = OwnedProcessCleanup.CaptureTrackedProcesses(pipeName);
        try
        {
            using var client = new ServiceClient(pipeName, connectTimeout: CommandTimeout, requestTimeout: CommandTimeout);
            var response = await client.SendAsync(new ServiceRequest { Command = "service.shutdown" }, cancellationToken);
            if (response.Success)
            {
                if (await WaitForDaemonExitAsync(pipeName, cancellationToken))
                {
                    return await WriteCleanupResultAsync(
                        pipeName,
                        preShutdownSnapshot,
                        cancellationToken,
                        "Service stopped.");
                }

                if (await TryForceStopTrackedDaemonAsync(
                    pipeName,
                    preShutdownSnapshot,
                    cancellationToken))
                {
                    Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = "Service stopped.", forced = true }, ServiceProtocol.JsonOptions));
                    return 0;
                }

                Console.WriteLine(JsonSerializer.Serialize(
                    new { success = false, error = $"Service acknowledged shutdown but did not exit within {ShutdownWaitTimeout.TotalSeconds:0} seconds." },
                    ServiceProtocol.JsonOptions));
                return 1;
            }

            if (!DaemonAutoStart.IsDaemonMutexHeld(pipeName))
            {
                return await WriteCleanupResultAsync(
                    pipeName,
                    preShutdownSnapshot,
                    cancellationToken,
                    "Service not running.");
            }

            if (await TryForceStopTrackedDaemonAsync(
                pipeName,
                preShutdownSnapshot,
                cancellationToken))
            {
                Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = "Service stopped.", forced = true }, ServiceProtocol.JsonOptions));
                return 0;
            }

            Console.WriteLine(JsonSerializer.Serialize(
                new
                {
                    success = false,
                    error = response.ErrorMessage ?? "Daemon is running but not responding, and the tracked daemon process could not be stopped."
                },
                ServiceProtocol.JsonOptions));
            return 1;
        }
        catch (Exception ex)
        {
            if (!DaemonAutoStart.IsDaemonMutexHeld(pipeName))
            {
                return await WriteCleanupResultAsync(
                    pipeName,
                    preShutdownSnapshot,
                    cancellationToken,
                    "Service not running.");
            }

            if (await TryForceStopTrackedDaemonAsync(
                pipeName,
                preShutdownSnapshot,
                cancellationToken))
            {
                Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = "Service stopped.", forced = true }, ServiceProtocol.JsonOptions));
                return 0;
            }

            Console.WriteLine(JsonSerializer.Serialize(
                new
                {
                    success = false,
                    error = $"Daemon is running but not responding, and the tracked daemon process could not be stopped. {ex.GetType().Name}: {ex.Message}"
                },
                ServiceProtocol.JsonOptions));
            return 1;
        }
    }

    private static async Task<bool> WaitForDaemonExitAsync(string pipeName, CancellationToken cancellationToken)
    {
        var deadline = DateTime.UtcNow + ShutdownWaitTimeout;
        while (DateTime.UtcNow < deadline)
        {
            if (!DaemonAutoStart.IsDaemonMutexHeld(pipeName))
            {
                return true;
            }

            await Task.Delay(ShutdownPollInterval, cancellationToken);
        }

        return !DaemonAutoStart.IsDaemonMutexHeld(pipeName);
    }

    private static async Task<bool> TryForceStopTrackedDaemonAsync(
        string pipeName,
        OwnedProcessCleanup.ProcessSnapshot preShutdownSnapshot,
        CancellationToken cancellationToken)
    {
        var cleanupResult = await OwnedProcessCleanup.CleanupAsync(
            pipeName,
            preShutdownSnapshot,
            cancellationToken);
        return cleanupResult.Success
            && cleanupResult.DaemonMatched
            && await WaitForDaemonExitAsync(pipeName, cancellationToken);
    }

    private static async Task<int> WriteCleanupResultAsync(
        string pipeName,
        OwnedProcessCleanup.ProcessSnapshot preShutdownSnapshot,
        CancellationToken cancellationToken,
        string successMessage)
    {
        var cleanupResult = await OwnedProcessCleanup.CleanupAsync(
            pipeName,
            preShutdownSnapshot,
            cancellationToken);
        if (!cleanupResult.Success)
        {
            Console.WriteLine(JsonSerializer.Serialize(
                new
                {
                    success = false,
                    error = cleanupResult.ErrorMessage
                        ?? $"One or more processes tracked for CLI pipe '{pipeName}' could not be stopped."
                },
                ServiceProtocol.JsonOptions));
            return 1;
        }

        Console.WriteLine(JsonSerializer.Serialize(
            new { success = true, message = successMessage },
            ServiceProtocol.JsonOptions));
        return 0;
    }
}

/// <summary>
/// Shows ExcelMCP CLI Service status including PID, session count, and uptime.
/// Surfaces actual error details instead of silently masking connection failures.
/// </summary>
internal sealed class ServiceStatusCommand : AsyncCommand
{
    protected override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
    {
        var pipeName = DaemonAutoStart.GetPipeName();
        var observation = DaemonConnectionPolicy.Observe(pipeName);
        var response = await DaemonConnectionPolicy.SendControlRequestAsync(
            pipeName,
            new ServiceRequest { Command = "service.status" },
            cancellationToken,
            observation.IsStopped
                ? DaemonConnectionPolicy.InitialProbeTimeout
                : DaemonConnectionPolicy.ControlTimeout);
        if (response.Success && response.Result != null)
        {
            var status = ServiceProtocol.Deserialize<ServiceStatus>(response.Result);
            if (status != null)
            {
                Console.WriteLine(JsonSerializer.Serialize(new
                {
                    success = true,
                    daemonState = DaemonConnectionPolicy.RunningState,
                    running = status.Running,
                    processId = status.ProcessId,
                    sessionCount = status.SessionCount,
                    startTime = status.StartTime,
                    uptime = status.Uptime.ToString(@"d\.hh\:mm\:ss", CultureInfo.InvariantCulture)
                }, ServiceProtocol.JsonOptions));
                return 0;
            }
        }

        if (response.Success)
        {
            response = new ServiceResponse
            {
                Success = false,
                Command = "service.status",
                ErrorCategory = "InvalidResponse",
                ErrorMessage = "Service returned an invalid status response."
            };
        }

        var failureState = DaemonConnectionPolicy.ResolveFailureState(pipeName, response);
        if (failureState.Name == DaemonConnectionPolicy.StoppedState)
        {
            return WriteStoppedStatus();
        }

        return CliErrorOutput.WriteDaemonError(response, failureState.Name, failureState.Running);
    }

    private static int WriteStoppedStatus()
    {
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            success = true,
            daemonState = DaemonConnectionPolicy.StoppedState,
            running = false,
            processId = 0,
            sessionCount = 0,
            startTime = (DateTime?)null,
            uptime = TimeSpan.Zero.ToString(@"d\.hh\:mm\:ss", CultureInfo.InvariantCulture)
        }, ServiceProtocol.JsonOptions));
        return 0;
    }
}
