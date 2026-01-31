using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Daemon;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Models.Actions;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// PivotTable commands - thin wrapper that sends requests to daemon.
/// Actions: list, read, create-from-range, create-from-table, create-from-datamodel, delete, refresh
/// </summary>
internal sealed class PivotTableCommand : AsyncCommand<PivotTableCommand.Settings>
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

        if (!ActionValidator.TryNormalizeAction<PivotTableAction>(settings.Action, out var action, out var errorMessage))
        {
            AnsiConsole.MarkupLine($"[red]{errorMessage}[/]");
            return 1;
        }
        var command = $"pivottable.{action}";

        // Note: property names must match daemon's Args classes (e.g., PivotTableFromRangeArgs)
        object? args = action switch
        {
            "list" => null,
            "read" => new { pivotTableName = settings.PivotTableName },
            "create-from-range" => new { pivotTableName = settings.PivotTableName, sourceSheet = settings.SourceSheet, sourceRange = settings.SourceRange, destinationSheet = settings.DestSheet, destinationCell = settings.DestCell },
            "create-from-table" => new { pivotTableName = settings.PivotTableName, tableName = settings.TableName, destinationSheet = settings.DestSheet, destinationCell = settings.DestCell },
            "create-from-datamodel" => new { pivotTableName = settings.PivotTableName, tableName = settings.TableName, destinationSheet = settings.DestSheet, destinationCell = settings.DestCell },
            "delete" => new { pivotTableName = settings.PivotTableName },
            "refresh" => new { pivotTableName = settings.PivotTableName },
            _ => new { pivotTableName = settings.PivotTableName }
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

        [CommandOption("--pivot-table <NAME>")]
        public string? PivotTableName { get; init; }

        [CommandOption("--table <NAME>")]
        public string? TableName { get; init; }

        [CommandOption("--source-sheet <NAME>")]
        public string? SourceSheet { get; init; }

        [CommandOption("--source-range <ADDRESS>")]
        public string? SourceRange { get; init; }

        [CommandOption("--dest-sheet <NAME>")]
        public string? DestSheet { get; init; }

        [CommandOption("--dest-cell <ADDRESS>")]
        public string? DestCell { get; init; }

        [CommandOption("--layout-style <STYLE>")]
        public string? LayoutStyle { get; init; }
    }
}
