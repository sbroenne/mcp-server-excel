using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
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
        // Check if daemon is already running
        var pipeName = Environment.GetEnvironmentVariable("EXCELMCP_CLI_PIPE") ?? ServiceSecurity.GetCliPipeName();
        using var checkClient = new ServiceClient(pipeName, connectTimeout: TimeSpan.FromSeconds(2));
        if (await checkClient.PingAsync(cancellationToken))
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = "Service already running." }, ServiceProtocol.JsonOptions));
            return 0;
        }

        // Launch daemon process
        var exePath = Environment.ProcessPath;
        if (string.IsNullOrEmpty(exePath))
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = "Cannot determine executable path." }, ServiceProtocol.JsonOptions));
            return 1;
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = exePath,
            Arguments = "service run",
            UseShellExecute = false,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden
        };

        try
        {
            Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = $"Failed to start daemon: {ex.Message}" }, ServiceProtocol.JsonOptions));
            return 1;
        }

        // Wait for daemon to be ready
        for (int i = 0; i < 20; i++)
        {
            await Task.Delay(250, cancellationToken);
            using var client = new ServiceClient(pipeName, connectTimeout: TimeSpan.FromSeconds(1));
            if (await client.PingAsync(cancellationToken))
            {
                Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = "Service started." }, ServiceProtocol.JsonOptions));
                return 0;
            }
        }

        Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = "Service started but not responding." }, ServiceProtocol.JsonOptions));
        return 1;
    }
}

/// <summary>
/// Gracefully stops the ExcelMCP CLI Service daemon.
/// </summary>
internal sealed class ServiceStopCommand : AsyncCommand
{
    public override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
    {
        var pipeName = Environment.GetEnvironmentVariable("EXCELMCP_CLI_PIPE") ?? ServiceSecurity.GetCliPipeName();
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
        var pipeName = Environment.GetEnvironmentVariable("EXCELMCP_CLI_PIPE") ?? ServiceSecurity.GetCliPipeName();
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
