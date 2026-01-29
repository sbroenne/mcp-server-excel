using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Daemon;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

// ============================================================================
// DAEMON COMMANDS
// ============================================================================

internal sealed class DaemonStartCommand : AsyncCommand
{
    public override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
    {
        if (await DaemonManager.IsDaemonRunningAsync(cancellationToken))
        {
            AnsiConsole.MarkupLine("[yellow]Daemon is already running.[/]");
            return 0;
        }

        AnsiConsole.MarkupLine("[dim]Starting daemon...[/]");
        var started = await DaemonManager.StartDaemonAsync(cancellationToken);

        if (started)
        {
            AnsiConsole.MarkupLine("[green]Daemon started successfully.[/]");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine("[red]Failed to start daemon.[/]");
            return 1;
        }
    }
}

internal sealed class DaemonStopCommand : AsyncCommand
{
    public override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
    {
        if (!await DaemonManager.IsDaemonRunningAsync(cancellationToken))
        {
            AnsiConsole.MarkupLine("[yellow]Daemon is not running.[/]");
            return 0;
        }

        var stopped = await DaemonManager.StopDaemonAsync(cancellationToken);

        if (stopped)
        {
            AnsiConsole.MarkupLine("[green]Daemon stopped.[/]");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine("[red]Failed to stop daemon.[/]");
            return 1;
        }
    }
}

internal sealed class DaemonStatusCommand : AsyncCommand
{
    public override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
    {
        var status = await DaemonManager.GetStatusAsync(cancellationToken);

        if (!status.Running)
        {
            AnsiConsole.MarkupLine("[yellow]Daemon is not running.[/]");
            Console.WriteLine(JsonSerializer.Serialize(new { running = false }, DaemonProtocol.JsonOptions));
            return 0;
        }

        var result = new
        {
            running = true,
            pid = status.ProcessId,
            sessions = status.SessionCount,
            uptime = status.Uptime.ToString(@"hh\:mm\:ss", CultureInfo.InvariantCulture)
        };

        Console.WriteLine(JsonSerializer.Serialize(result, DaemonProtocol.JsonOptions));
        return 0;
    }
}

// ============================================================================
// SESSION COMMANDS
// ============================================================================

internal sealed class SessionOpenCommand : AsyncCommand<SessionOpenCommand.Settings>
{
    public override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.FilePath))
        {
            AnsiConsole.MarkupLine("[red]File path is required.[/]");
            return 1;
        }

        // Ensure daemon is running
        if (!await DaemonManager.EnsureDaemonRunningAsync(cancellationToken))
        {
            AnsiConsole.MarkupLine("[red]Failed to start daemon.[/]");
            return 1;
        }

        using var client = new DaemonClient();
        var response = await client.SendAsync(new DaemonRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new { filePath = settings.FilePath }, DaemonProtocol.JsonOptions)
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(response.Result);
            return 0;
        }
        else
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, DaemonProtocol.JsonOptions));
            return 1;
        }
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<FILE>")]
        public string FilePath { get; init; } = string.Empty;
    }
}

internal sealed class SessionCloseCommand : AsyncCommand<SessionCloseCommand.Settings>
{
    public override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.SessionId))
        {
            AnsiConsole.MarkupLine("[red]Session ID is required.[/]");
            return 1;
        }

        using var client = new DaemonClient();
        var response = await client.SendAsync(new DaemonRequest
        {
            Command = "session.close",
            SessionId = settings.SessionId,
            Args = JsonSerializer.Serialize(new { save = settings.Save }, DaemonProtocol.JsonOptions)
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = settings.Save ? "Session closed and saved." : "Session closed." }, DaemonProtocol.JsonOptions));
            return 0;
        }
        else
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, DaemonProtocol.JsonOptions));
            return 1;
        }
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandOption("-s|--session <SESSION>")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--save")]
        public bool Save { get; init; }
    }
}

internal sealed class SessionListCommand : AsyncCommand
{
    public override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
    {
        if (!await DaemonManager.IsDaemonRunningAsync(cancellationToken))
        {
            Console.WriteLine(JsonSerializer.Serialize(new { sessions = Array.Empty<object>() }, DaemonProtocol.JsonOptions));
            return 0;
        }

        using var client = new DaemonClient();
        var response = await client.SendAsync(new DaemonRequest { Command = "session.list" }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(response.Result);
            return 0;
        }
        else
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, DaemonProtocol.JsonOptions));
            return 1;
        }
    }
}

internal sealed class SessionSaveCommand : AsyncCommand<SessionSaveCommand.Settings>
{
    public override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.SessionId))
        {
            AnsiConsole.MarkupLine("[red]Session ID is required.[/]");
            return 1;
        }

        using var client = new DaemonClient();
        var response = await client.SendAsync(new DaemonRequest
        {
            Command = "session.save",
            SessionId = settings.SessionId
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = "Session saved." }, DaemonProtocol.JsonOptions));
            return 0;
        }
        else
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, DaemonProtocol.JsonOptions));
            return 1;
        }
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandOption("-s|--session <SESSION>")]
        public string SessionId { get; init; } = string.Empty;
    }
}
