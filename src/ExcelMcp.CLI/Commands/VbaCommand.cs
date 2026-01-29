using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Daemon;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// VBA commands - thin wrapper that sends requests to daemon.
/// Actions: list, view, import, delete, run, update
/// </summary>
internal sealed class VbaCommand : AsyncCommand<VbaCommand.Settings>
{
    public override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.SessionId))
        {
            AnsiConsole.MarkupLine("[red]Session ID is required. Use --session <id>[/]");
            return 1;
        }

        if (string.IsNullOrWhiteSpace(settings.Action))
        {
            AnsiConsole.MarkupLine("[red]Action is required.[/]");
            return 1;
        }

        var action = settings.Action.Trim().ToLowerInvariant();
        var command = $"vba.{action}";

        object? args = action switch
        {
            "list" => null,
            "view" => new { moduleName = settings.ModuleName },
            "import" => new { moduleName = settings.ModuleName, code = settings.Code, moduleType = settings.ModuleType },
            "delete" => new { moduleName = settings.ModuleName },
            "run" => new { macroName = settings.MacroName, arguments = settings.Arguments },
            "update" => new { moduleName = settings.ModuleName, code = settings.Code },
            _ => new { moduleName = settings.ModuleName }
        };

        using var client = new DaemonClient();
        var response = await client.SendAsync(new DaemonRequest
        {
            Command = command,
            SessionId = settings.SessionId,
            Args = args != null ? JsonSerializer.Serialize(args, DaemonProtocol.JsonOptions) : null
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(!string.IsNullOrEmpty(response.Result) ? response.Result : JsonSerializer.Serialize(new { success = true }, DaemonProtocol.JsonOptions));
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
        [CommandArgument(0, "<ACTION>")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--module <NAME>")]
        public string? ModuleName { get; init; }

        [CommandOption("--macro <NAME>")]
        public string? MacroName { get; init; }

        [CommandOption("--code <CODE>")]
        public string? Code { get; init; }

        [CommandOption("--module-type <TYPE>")]
        public string? ModuleType { get; init; }

        [CommandOption("--arguments <ARGS>")]
        public string? Arguments { get; init; }
    }
}
