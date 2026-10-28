using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Telemetry;
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
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "CLI service-start boundary writes success=false JSON with the startup error message and returns exit code 1.")]
    public override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
    {
        try
        {
            using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
            return CliResultOutput.Write(new ServiceResponse { Success = true });
        }
        catch (Exception ex)
        {
            return CliErrorOutput.WriteException(ex);
        }
    }
}

/// <summary>
/// Requests normal CLI service shutdown without terminating retained Excel processes.
/// </summary>
internal sealed class ServiceStopCommand : AsyncCommand
{
    private static readonly TimeSpan CommandTimeout = TimeSpan.FromSeconds(2);
    internal static readonly TimeSpan ShutdownRequestTimeout = TimeSpan.FromSeconds(10);
    private static readonly TimeSpan ShutdownWaitTimeout = TimeSpan.FromSeconds(10);
    private static readonly TimeSpan ShutdownPollInterval = TimeSpan.FromMilliseconds(250);

    public override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
    {
        var pipeName = DaemonAutoStart.GetPipeName();
        return await DaemonAutoStart.WithStartupLockAsync(
            pipeName,
            () => ExecuteWithStartupLockAsync(pipeName, cancellationToken),
            cancellationToken);
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "CLI shutdown boundary reports unconfirmed shutdown or cleanup errors as success=false JSON and returns exit code 1 without terminating a running service.")]
    private static async Task<int> ExecuteWithStartupLockAsync(
        string pipeName,
        CancellationToken cancellationToken)
    {
        var preShutdownSnapshot = OwnedProcessCleanup.CaptureTrackedProcesses(pipeName);
        try
        {
            using var client = new ServiceClient(pipeName, connectTimeout: CommandTimeout,
                requestTimeout: ShutdownRequestTimeout, useControlEndpoint: true);
            var request = new ServiceRequest { Command = "service.shutdown" };
            var response = await CliTelemetry.TrackCommandAsync(
                request,
                // A missing pipe would make the connect wait out its whole timeout.
                () => !DaemonAutoStart.IsDaemonMutexHeld(pipeName) && !DaemonAutoStart.ControlPipeExists(pipeName)
                    ? Task.FromResult(new ServiceResponse
                    {
                        Success = false,
                        Command = request.Command,
                        ErrorCategory = "ServiceUnavailable",
                        ErrorMessage = "Cannot connect to service. Is it running?"
                    })
                    : client.SendAsync(request, ShutdownRequestTimeout, cancellationToken));
            if (!response.Success
                && !client.HasConnected
                && DaemonConnectionPolicy.IsTransportFailure(response)
                && !DaemonAutoStart.IsDaemonMutexHeld(pipeName))
            {
                return await WriteCleanupResultAsync(
                    pipeName,
                    preShutdownSnapshot,
                    cancellationToken,
                    "Service not running.");
            }

            if (response.Success)
            {
                if (await WaitForDaemonExitAsync(pipeName, cancellationToken))
                {
                    return await WriteCleanupResultAsync(
                        pipeName,
                        preShutdownSnapshot,
                        cancellationToken,
                        "Service stopped.",
                        response.PartialState);
                }

                CliResultOutput.WriteJson(JsonSerializer.Serialize(
                    new
                    {
                        success = false,
                        errorCategory = "Timeout",
                        errorMessage = $"Service acknowledged shutdown but did not exit within {ShutdownWaitTimeout.TotalSeconds:0} seconds. " +
                            "The service and its workbooks have not been terminated."
                    },
                    ServiceProtocol.JsonOptions));
                return 1;
            }

            return CliResultOutput.Write(response);
        }
        catch (Exception ex) when (ex is OperationCanceledException or TimeoutException)
        {
            CliResultOutput.WriteJson(JsonSerializer.Serialize(
                new
                {
                    success = false,
                    errorCategory = ex is TimeoutException ? "Timeout" : "Cancelled",
                    errorMessage = "The shutdown wait ended without confirmation. The service and its workbooks have not been terminated.",
                    cause = ex.Message
                }, ServiceProtocol.JsonOptions));
            return 1;
        }
        catch (Exception ex)
        {
            CliResultOutput.WriteJson(JsonSerializer.Serialize(
                new
                {
                    success = false,
                    errorMessage = "The shutdown response could not be confirmed. The service and its workbooks have not been terminated. " +
                        $"{ex.GetType().Name}: {ex.Message}"
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

    private static async Task<int> WriteCleanupResultAsync(
        string pipeName,
        OwnedProcessCleanup.ProcessSnapshot preShutdownSnapshot,
        CancellationToken cancellationToken,
        string successMessage,
        string? shutdownWarning = null)
    {
        var cleanupResult = await OwnedProcessCleanup.CleanupAsync(
            pipeName,
            preShutdownSnapshot,
            cancellationToken);
        if (!cleanupResult.Success)
        {
            CliResultOutput.WriteJson(JsonSerializer.Serialize(
                new
                {
                    success = false,
                    errorMessage = cleanupResult.ErrorMessage
                        ?? "Recorded Excel copies could not be checked. Ownership checks did not close Excel."
                },
                ServiceProtocol.JsonOptions));
            return 1;
        }

        var retained = OwnedProcessCleanup.GetRetainedExcelProcesses(pipeName, serviceRunning: false);
        CliResultOutput.WriteJson(JsonSerializer.Serialize(
            new
            {
                success = true,
                message = successMessage,
                shutdownWarning,
                retainedExcelProcessCount = retained.Count,
                retainedExcelProcesses = retained,
                recoveryMessage = retained.Count > 0 ? OwnedProcessCleanup.RecoveryMessage : null
            },
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
    public override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
    {
        var pipeName = DaemonAutoStart.GetPipeName();
        var request = new ServiceRequest { Command = "service.status" };
        var response = await CliTelemetry.TrackCommandAsync(
            request,
            () => DaemonConnectionPolicy.SendObservedControlRequestAsync(
                pipeName,
                request,
                cancellationToken));
        if (response.Success && response.Result != null)
        {
            ServiceStatus? status;
            try
            {
                status = ServiceProtocol.Deserialize<ServiceStatus>(response.Result);
            }
            catch (JsonException)
            {
                CliTelemetry.RecordFinalFailure("InvalidResponse");
                throw;
            }
            if (status != null)
            {
                var (retained, retainedError) = ReadRetainedExcelProcesses(pipeName, serviceRunning: true);
                CliResultOutput.WriteJson(JsonSerializer.Serialize(new
                {
                    success = retainedError is null,
                    daemonState = DaemonConnectionPolicy.RunningState,
                    running = status.Running,
                    processId = status.ProcessId,
                    sessionCount = status.SessionCount,
                    startTime = status.StartTime,
                    uptime = status.Uptime.ToString(@"d\.hh\:mm\:ss", CultureInfo.InvariantCulture),
                    retainedExcelProcessCount = retained?.Count,
                    retainedExcelProcesses = retained,
                    recoveryMessage = retained?.Count > 0 ? OwnedProcessCleanup.RecoveryMessage : null,
                    retainedExcelError = retainedError,
                    errorMessage = retainedError
                }, ServiceProtocol.JsonOptions));
                return retainedError is null ? 0 : 1;
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
            CliTelemetry.RecordFinalFailure(response.ErrorCategory);
        }

        var failureState = DaemonConnectionPolicy.ResolveFailureState(pipeName, response);
        if (failureState.Name == DaemonConnectionPolicy.StoppedState)
        {
            return WriteStoppedStatus(pipeName);
        }

        var (retainedProcesses, retainedErrorMessage) = ReadRetainedExcelProcesses(pipeName, failureState.Running);
        return CliErrorOutput.WriteDaemonError(response, failureState.Name, failureState.Running,
            retainedProcesses, retainedErrorMessage);
    }

    private static int WriteStoppedStatus(string pipeName)
    {
        var (retained, retainedError) = ReadRetainedExcelProcesses(pipeName, serviceRunning: false);
        CliResultOutput.WriteJson(JsonSerializer.Serialize(new
        {
            success = retainedError is null,
            daemonState = DaemonConnectionPolicy.StoppedState,
            running = false,
            processId = 0,
            sessionCount = 0,
            startTime = (DateTime?)null,
            uptime = TimeSpan.Zero.ToString(@"d\.hh\:mm\:ss", CultureInfo.InvariantCulture),
            retainedExcelProcessCount = retained?.Count,
            retainedExcelProcesses = retained,
            recoveryMessage = retained?.Count > 0 ? OwnedProcessCleanup.RecoveryMessage : null,
            retainedExcelError = retainedError,
            errorMessage = retainedError
        }, ServiceProtocol.JsonOptions));
        return retainedError is null ? 0 : 1;
    }

    private static (IReadOnlyList<OwnedProcessCleanup.RetainedExcelProcess>? Processes, string? Error)
        ReadRetainedExcelProcesses(string pipeName, bool serviceRunning)
    {
        try
        {
            return (OwnedProcessCleanup.GetRetainedExcelProcesses(pipeName, serviceRunning), null);
        }
        catch (InvalidOperationException ex)
        {
            return (null, ex.Message);
        }
    }
}
