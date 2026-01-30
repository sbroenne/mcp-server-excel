using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Daemon;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Models.Actions;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// PowerQuery commands - thin wrapper that sends requests to daemon.
/// Actions: list, view, create, update, rename, delete, refresh, refresh-all, load-to, get-load-config
/// </summary>
internal sealed class PowerQueryCommand : AsyncCommand<PowerQueryCommand.Settings>
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

        if (!ActionValidator.TryNormalizeAction<PowerQueryAction>(settings.Action, out var action, out var errorMessage))
        {
            AnsiConsole.MarkupLine($"[red]{errorMessage}[/]");
            return 1;
        }
        var command = $"powerquery.{action}";

        // Note: property names must match daemon's Args classes (e.g., PowerQueryRenameArgs)
        object? args = action switch
        {
            "list" => null,
            "view" => new { queryName = settings.QueryName },
            "create" => new { queryName = settings.QueryName, mCode = settings.MCode, loadDestination = settings.LoadDestination },
            "update" => new { queryName = settings.QueryName, mCode = settings.MCode },
            "rename" => new { oldName = settings.QueryName, newName = settings.NewName },
            "delete" => new { queryName = settings.QueryName },
            "refresh" => new { queryName = settings.QueryName },
            "refresh-all" => null,
            "load-to" => new { queryName = settings.QueryName, loadDestination = settings.LoadDestination },
            "get-load-config" => new { queryName = settings.QueryName },
            _ => new { queryName = settings.QueryName }
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

        [CommandOption("--query <NAME>")]
        public string? QueryName { get; init; }

        [CommandOption("--new-name <NAME>")]
        public string? NewName { get; init; }

        [CommandOption("--mcode <CODE>")]
        public string? MCode { get; init; }

        [CommandOption("--load-destination <DEST>")]
        public string? LoadDestination { get; init; }
    }
}
