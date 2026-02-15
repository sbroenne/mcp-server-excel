using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

// ============================================================================
// SERVICE LIFECYCLE COMMANDS
// ============================================================================

/// <summary>
/// Starts the ExcelMCP Service if not already running.
/// </summary>
internal sealed class ServiceStartCommand : AsyncCommand
{
    public override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
    {
        var started = await ServiceManager.EnsureServiceRunningAsync(cancellationToken);
        if (started)
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = "Service started." }, ServiceProtocol.JsonOptions));
            return 0;
        }
        else
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = "Failed to start service." }, ServiceProtocol.JsonOptions));
            return 1;
        }
    }
}

/// <summary>
/// Gracefully stops the ExcelMCP Service.
/// </summary>
internal sealed class ServiceStopCommand : AsyncCommand
{
    public override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
    {
        var stopped = await ServiceManager.StopServiceAsync(cancellationToken);
        if (stopped)
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = "Service stopped." }, ServiceProtocol.JsonOptions));
            return 0;
        }
        else
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = "Failed to stop service." }, ServiceProtocol.JsonOptions));
            return 1;
        }
    }
}

/// <summary>
/// Shows ExcelMCP Service status including PID, session count, and uptime.
/// </summary>
internal sealed class ServiceStatusCommand : AsyncCommand
{
    public override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
    {
        var status = await ServiceManager.GetStatusAsync(cancellationToken);

        Console.WriteLine(JsonSerializer.Serialize(new
        {
            success = true,
            running = status.Running,
            processId = status.ProcessId,
            sessionCount = status.SessionCount,
            startTime = status.Running ? status.StartTime : (DateTime?)null,
            uptime = status.Running ? status.Uptime.ToString(@"d\.hh\:mm\:ss", CultureInfo.InvariantCulture) : null
        }, ServiceProtocol.JsonOptions));

        return 0;
    }
}
