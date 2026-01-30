using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Daemon;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// NamedRange commands - thin wrapper that sends requests to daemon.
/// Actions: list, read, write, create, update, delete
/// </summary>
internal sealed class NamedRangeCommand : AsyncCommand<NamedRangeCommand.Settings>
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
        var command = $"namedrange.{action}";

        // Note: property names must match daemon's Args classes (e.g., NamedRangeArgs, NamedRangeCreateArgs)
        object? args = action switch
        {
            "list" => null,
            "read" => new { paramName = settings.Name },
            "write" => new { paramName = settings.Name, value = settings.Value },
            "create" => new { paramName = settings.Name, reference = settings.RefersTo },
            "update" => new { paramName = settings.Name, reference = settings.RefersTo },
            "delete" => new { paramName = settings.Name },
            _ => new { paramName = settings.Name }
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

        [CommandOption("--name <NAME>")]
        public string? Name { get; init; }

        [CommandOption("--refers-to <FORMULA>")]
        public string? RefersTo { get; init; }

        [CommandOption("--value <VALUE>")]
        public string? Value { get; init; }

        [CommandOption("--sheet-scope <SHEET>")]
        public string? SheetScope { get; init; }
    }
}
