using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Daemon;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Models.Actions;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Connection commands - thin wrapper that sends requests to daemon.
/// Actions: list, view, create, test, refresh, delete, load-to, get-properties, set-properties
/// </summary>
internal sealed class ConnectionCommand : AsyncCommand<ConnectionCommand.Settings>
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

        if (!ActionValidator.TryNormalizeAction<ConnectionAction>(settings.Action, out var action, out var errorMessage))
        {
            AnsiConsole.MarkupLine($"[red]{errorMessage}[/]");
            return 1;
        }
        var command = $"connection.{action}";

        object? args = action switch
        {
            "list" => null,
            "view" => new { connectionName = settings.ConnectionName },
            "create" => new { connectionName = settings.ConnectionName, connectionType = settings.ConnectionType, connectionString = settings.ConnectionString, commandText = settings.CommandText },
            "test" => new { connectionName = settings.ConnectionName },
            "refresh" => new { connectionName = settings.ConnectionName },
            "delete" => new { connectionName = settings.ConnectionName },
            "load-to" => new { connectionName = settings.ConnectionName, loadDestination = settings.LoadDestination, sheetName = settings.SheetName, targetCell = settings.TargetCell },
            "get-properties" => new { connectionName = settings.ConnectionName },
            "set-properties" => new { connectionName = settings.ConnectionName, refreshOnOpen = settings.RefreshOnOpen, enableRefresh = settings.EnableRefresh },
            _ => new { connectionName = settings.ConnectionName }
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

        [CommandOption("--connection <NAME>")]
        public string? ConnectionName { get; init; }

        [CommandOption("--connection-type <TYPE>")]
        public string? ConnectionType { get; init; }

        [CommandOption("--connection-string <STRING>")]
        public string? ConnectionString { get; init; }

        [CommandOption("--command-text <TEXT>")]
        public string? CommandText { get; init; }

        [CommandOption("--load-destination <DEST>")]
        public string? LoadDestination { get; init; }

        [CommandOption("--sheet <NAME>")]
        public string? SheetName { get; init; }

        [CommandOption("--target-cell <ADDRESS>")]
        public string? TargetCell { get; init; }

        [CommandOption("--refresh-on-open")]
        public bool? RefreshOnOpen { get; init; }

        [CommandOption("--enable-refresh")]
        public bool? EnableRefresh { get; init; }
    }
}
