using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Daemon;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Models.Actions;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Chart commands - thin wrapper that sends requests to daemon.
/// Actions: list, read, create-from-range, create-from-pivottable, delete, move, fit-to-range
/// </summary>
internal sealed class ChartCommand : AsyncCommand<ChartCommand.Settings>
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

        if (!ActionValidator.TryNormalizeAction<ChartAction>(settings.Action, out var action, out var errorMessage))
        {
            AnsiConsole.MarkupLine($"[red]{errorMessage}[/]");
            return 1;
        }
        var command = $"chart.{action}";

        // Note: property names must match daemon's Args classes (e.g., ChartFromRangeArgs)
        object? args = action switch
        {
            "list" => new { sheetName = settings.SheetName },
            "read" => new { sheetName = settings.SheetName, chartName = settings.ChartName },
            "create-from-range" => new { sheetName = settings.SheetName, chartName = settings.ChartName, sourceRange = settings.SourceRange, chartType = settings.ChartType },
            "create-from-pivottable" => new { pivotTableName = settings.PivotTableName, sheetName = settings.SheetName, chartName = settings.ChartName, chartType = settings.ChartType },
            "delete" => new { sheetName = settings.SheetName, chartName = settings.ChartName },
            "move" => new { chartName = settings.ChartName },
            "fit-to-range" => new { chartName = settings.ChartName, sheetName = settings.SheetName, rangeAddress = settings.TargetRange },
            _ => new { sheetName = settings.SheetName, chartName = settings.ChartName }
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

        [CommandOption("--sheet <NAME>")]
        public string? SheetName { get; init; }

        [CommandOption("--chart <NAME>")]
        public string? ChartName { get; init; }

        [CommandOption("--pivot-table <NAME>")]
        public string? PivotTableName { get; init; }

        [CommandOption("--source-range <ADDRESS>")]
        public string? SourceRange { get; init; }

        [CommandOption("--target-sheet <NAME>")]
        public string? TargetSheet { get; init; }

        [CommandOption("--target-range <ADDRESS>")]
        public string? TargetRange { get; init; }

        [CommandOption("--chart-type <TYPE>")]
        public string? ChartType { get; init; }

        [CommandOption("--position <POSITION>")]
        public string? Position { get; init; }
    }
}
