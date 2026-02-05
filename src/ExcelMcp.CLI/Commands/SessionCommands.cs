using System.ComponentModel;
using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

// ============================================================================
// SESSION COMMANDS
// ============================================================================

internal sealed class SessionCreateCommand : AsyncCommand<SessionCreateCommand.Settings>
{
    public override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.FilePath))
        {
            AnsiConsole.MarkupLine("[red]File path is required.[/]");
            return 1;
        }

        // Ensure service is running
        if (!await ServiceManager.EnsureServiceRunningAsync(cancellationToken))
        {
            AnsiConsole.MarkupLine("[red]Failed to start service.[/]");
            return 1;
        }

        using var client = new ServiceClient();
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = "session.create",
            Args = JsonSerializer.Serialize(new { filePath = settings.FilePath, timeoutSeconds = settings.TimeoutSeconds }, ServiceProtocol.JsonOptions)
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(response.Result);
            return 0;
        }
        else
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, ServiceProtocol.JsonOptions));
            return 1;
        }
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<FILE>")]
        [Description("Path to the new Excel file to create")]
        public string FilePath { get; init; } = string.Empty;

        [CommandOption("--timeout <SECONDS>")]
        [Description("Session timeout in seconds")]
        public int? TimeoutSeconds { get; init; }
    }
}

internal sealed class SessionOpenCommand : AsyncCommand<SessionOpenCommand.Settings>
{
    public override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.FilePath))
        {
            AnsiConsole.MarkupLine("[red]File path is required.[/]");
            return 1;
        }

        // Ensure service is running
        if (!await ServiceManager.EnsureServiceRunningAsync(cancellationToken))
        {
            AnsiConsole.MarkupLine("[red]Failed to start service.[/]");
            return 1;
        }

        using var client = new ServiceClient();
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new { filePath = settings.FilePath, timeoutSeconds = settings.TimeoutSeconds }, ServiceProtocol.JsonOptions)
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(response.Result);
            return 0;
        }
        else
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, ServiceProtocol.JsonOptions));
            return 1;
        }
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<FILE>")]
        [Description("Path to the Excel file to open")]
        public string FilePath { get; init; } = string.Empty;

        [CommandOption("--timeout <SECONDS>")]
        [Description("Session timeout in seconds")]
        public int? TimeoutSeconds { get; init; }
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

        using var client = new ServiceClient();
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = settings.SessionId,
            Args = JsonSerializer.Serialize(new { save = settings.Save }, ServiceProtocol.JsonOptions)
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = settings.Save ? "Session closed and saved." : "Session closed." }, ServiceProtocol.JsonOptions));
            return 0;
        }
        else
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, ServiceProtocol.JsonOptions));
            return 1;
        }
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandOption("-s|--session <SESSION>")]
        [Description("Session ID to close")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--save")]
        [Description("Save changes before closing")]
        public bool Save { get; init; }
    }
}

internal sealed class SessionListCommand : AsyncCommand
{
    public override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
    {
        if (!await ServiceManager.IsServiceRunningAsync(cancellationToken))
        {
            Console.WriteLine(JsonSerializer.Serialize(new { sessions = Array.Empty<object>() }, ServiceProtocol.JsonOptions));
            return 0;
        }

        using var client = new ServiceClient();
        var response = await client.SendAsync(new ServiceRequest { Command = "session.list" }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(response.Result);
            return 0;
        }
        else
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, ServiceProtocol.JsonOptions));
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

        using var client = new ServiceClient();
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = "session.save",
            SessionId = settings.SessionId
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = "Session saved." }, ServiceProtocol.JsonOptions));
            return 0;
        }
        else
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, ServiceProtocol.JsonOptions));
            return 1;
        }
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandOption("-s|--session <SESSION>")]
        [Description("Session ID to save")]
        public string SessionId { get; init; } = string.Empty;
    }
}



