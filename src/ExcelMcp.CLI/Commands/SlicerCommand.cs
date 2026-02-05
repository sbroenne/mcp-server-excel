using System.ComponentModel;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Service;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Models.Actions;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Slicer commands - thin wrapper that sends requests to service.
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

        using var client = new ServiceClient();
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = command,
            SessionId = settings.SessionId,
            Args = args != null ? JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions) : null
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(!string.IsNullOrEmpty(response.Result) ? response.Result : JsonSerializer.Serialize(new { success = true }, ServiceProtocol.JsonOptions));
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
        [CommandArgument(0, "<ACTION>")]
        [Description("The action to perform (e.g., create-slicer, list-slicers, set-slicer-selection)")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        [Description("Session ID from 'session open' command")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--slicer <NAME>")]
        [Description("Slicer name")]
        public string? SlicerName { get; init; }

        [CommandOption("--pivottable <NAME>")]
        [Description("PivotTable name for slicer source")]
        public string? PivotTableName { get; init; }

        [CommandOption("--table <NAME>")]
        [Description("Table name for table slicer source")]
        public string? TableName { get; init; }

        [CommandOption("--source-field <NAME>")]
        [Description("PivotTable field name for slicer")]
        public string? SourceFieldName { get; init; }

        [CommandOption("--column <NAME>")]
        [Description("Table column name for table slicer")]
        public string? ColumnName { get; init; }

        [CommandOption("--destination-sheet <NAME>")]
        [Description("Worksheet for slicer placement")]
        public string? DestinationSheet { get; init; }

        [CommandOption("--top <VALUE>")]
        [Description("Top position in points")]
        public double? Top { get; init; }

        [CommandOption("--left <VALUE>")]
        [Description("Left position in points")]
        public double? Left { get; init; }

        [CommandOption("--width <VALUE>")]
        [Description("Width in points")]
        public double? Width { get; init; }

        [CommandOption("--height <VALUE>")]
        [Description("Height in points")]
        public double? Height { get; init; }

        [CommandOption("--selected-items <ITEMS>")]
        [Description("Comma-separated items to select")]
        public string? SelectedItems { get; init; }

        [CommandOption("--multi-select")]
        [Description("Allow multiple item selection")]
        public bool? MultiSelect { get; init; }

        [CommandOption("--target-pivottable <NAME>")]
        [Description("Target PivotTable for slicer connection")]
        public string? TargetPivotTableName { get; init; }
    }
}
