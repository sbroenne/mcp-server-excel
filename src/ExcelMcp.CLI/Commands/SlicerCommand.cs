using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Daemon;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Models.Actions;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Slicer commands - thin wrapper that sends requests to daemon.
/// Actions: create-slicer, list-slicers, set-slicer-selection, delete-slicer,
/// create-table-slicer, list-table-slicers, set-table-slicer-selection, delete-table-slicer
/// </summary>
internal sealed class SlicerCommand : AsyncCommand<SlicerCommand.Settings>
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

        if (!ActionValidator.TryNormalizeAction<SlicerAction>(settings.Action, out var action, out var errorMessage))
        {
            AnsiConsole.MarkupLine($"[red]{errorMessage}[/]");
            return 1;
        }
        var command = $"slicer.{action}";

        object? args = action switch
        {
            "list-slicers" => new { pivotTableName = settings.PivotTableName },
            "list-table-slicers" => new { tableName = settings.TableName },
            "create-slicer" => new
            {
                pivotTableName = settings.PivotTableName,
                sourceFieldName = settings.SourceFieldName,
                slicerName = settings.SlicerName,
                destinationSheet = settings.DestinationSheet,
                top = settings.Top,
                left = settings.Left,
                width = settings.Width,
                height = settings.Height
            },
            "create-table-slicer" => new
            {
                tableName = settings.TableName,
                columnName = settings.ColumnName,
                slicerName = settings.SlicerName,
                destinationSheet = settings.DestinationSheet,
                top = settings.Top,
                left = settings.Left,
                width = settings.Width,
                height = settings.Height
            },
            "set-slicer-selection" => new { slicerName = settings.SlicerName, selectedItems = settings.SelectedItems, multiSelect = settings.MultiSelect },
            "set-table-slicer-selection" => new { slicerName = settings.SlicerName, selectedItems = settings.SelectedItems, multiSelect = settings.MultiSelect },
            "delete-slicer" => new { slicerName = settings.SlicerName },
            "delete-table-slicer" => new { slicerName = settings.SlicerName },
            _ => new { slicerName = settings.SlicerName }
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

        [CommandOption("--slicer <NAME>")]
        public string? SlicerName { get; init; }

        [CommandOption("--pivottable <NAME>")]
        public string? PivotTableName { get; init; }

        [CommandOption("--table <NAME>")]
        public string? TableName { get; init; }

        [CommandOption("--source-field <NAME>")]
        public string? SourceFieldName { get; init; }

        [CommandOption("--column <NAME>")]
        public string? ColumnName { get; init; }

        [CommandOption("--destination-sheet <NAME>")]
        public string? DestinationSheet { get; init; }

        [CommandOption("--top <VALUE>")]
        public double? Top { get; init; }

        [CommandOption("--left <VALUE>")]
        public double? Left { get; init; }

        [CommandOption("--width <VALUE>")]
        public double? Width { get; init; }

        [CommandOption("--height <VALUE>")]
        public double? Height { get; init; }

        [CommandOption("--selected-items <ITEMS>")]
        public string? SelectedItems { get; init; }

        [CommandOption("--multi-select")]
        public bool? MultiSelect { get; init; }

        [CommandOption("--target-pivottable <NAME>")]
        public string? TargetPivotTableName { get; init; }
    }
}
