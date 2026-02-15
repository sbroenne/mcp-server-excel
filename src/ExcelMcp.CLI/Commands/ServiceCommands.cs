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
    public override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
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
/// Gracefully stops the ExcelMCP CLI Service daemon.
/// </summary>
internal sealed class ServiceStopCommand : AsyncCommand
{
    public override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
    {
        var pipeName = DaemonAutoStart.GetPipeName();
        try
        {
            using var client = new ServiceClient(pipeName, connectTimeout: TimeSpan.FromSeconds(2));
            var response = await client.SendAsync(new ServiceRequest { Command = "service.shutdown" }, cancellationToken);
            if (response.Success)
            {
                Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = "Service stopped." }, ServiceProtocol.JsonOptions));
                return 0;
            }
            else
            {
                Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage ?? "Failed to stop service." }, ServiceProtocol.JsonOptions));
                return 1;
            }
        }
        catch (Exception)
        {
            // Can't connect — daemon not running
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = "Service not running." }, ServiceProtocol.JsonOptions));
            return 0;
        }
    }
}

/// <summary>
/// Shows ExcelMCP CLI Service status including PID, session count, and uptime.
/// </summary>
internal sealed class ServiceStatusCommand : AsyncCommand
{
    public override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
    {
        var pipeName = DaemonAutoStart.GetPipeName();
        try
        {
            using var client = new ServiceClient(pipeName, connectTimeout: TimeSpan.FromSeconds(2));
            var response = await client.SendAsync(new ServiceRequest { Command = "service.status" }, cancellationToken);
            if (response.Success && response.Result != null)
            {
                var status = ServiceProtocol.Deserialize<ServiceStatus>(response.Result);
                if (status != null)
                {
                    Console.WriteLine(JsonSerializer.Serialize(new
                    {
                        success = true,
                        running = status.Running,
                        processId = status.ProcessId,
                        sessionCount = status.SessionCount,
                        startTime = status.StartTime,
                        uptime = status.Uptime.ToString(@"d\.hh\:mm\:ss", CultureInfo.InvariantCulture)
                    }, ServiceProtocol.JsonOptions));
                    return 0;
                }
            }
        }
        catch (Exception)
        {
            // Can't connect — daemon not running
        }

        Console.WriteLine(JsonSerializer.Serialize(new
        {
            success = true,
            running = false,
            processId = 0,
            sessionCount = 0
        }, ServiceProtocol.JsonOptions));
        return 0;
    }
}
