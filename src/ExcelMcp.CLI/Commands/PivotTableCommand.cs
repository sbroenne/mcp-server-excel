using System.ComponentModel;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Service;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Models.Actions;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// PivotTable commands - thin wrapper that sends requests to service.
/// Actions: list, read, create-from-range, create-from-table, create-from-datamodel, delete, refresh
/// Plus field actions: list-fields, add-row-field, add-column-field, add-value-field, add-filter-field, remove-field, etc.
/// Plus calc actions: set-layout, set-subtotals, set-grand-totals, get-data, calculated fields/members
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

        // Combine all PivotTable action types
        var validActions = ActionValidator.GetValidActions<PivotTableAction>()
            .Concat(ActionValidator.GetValidActions<PivotTableFieldAction>())
            .Concat(ActionValidator.GetValidActions<PivotTableCalcAction>())
            .ToArray();

        if (!ActionValidator.TryNormalizeAction(settings.Action, validActions, out var action, out var errorMessage))
        {
            AnsiConsole.MarkupLine($"[red]{errorMessage}[/]");
            return 1;
        }
        var command = $"pivottable.{action}";

        // Note: property names must match daemon's Args classes (e.g., PivotTableFromRangeArgs)
        object? args = action switch
        {
            // PivotTableAction
            "list" => null,
            "read" => new { pivotTableName = settings.PivotTableName },
            "create-from-range" => new { pivotTableName = settings.PivotTableName, sourceSheet = settings.SourceSheet, sourceRange = settings.SourceRange, destinationSheet = settings.DestSheet, destinationCell = settings.DestCell },
            "create-from-table" => new { pivotTableName = settings.PivotTableName, tableName = settings.TableName, destinationSheet = settings.DestSheet, destinationCell = settings.DestCell },
            "create-from-datamodel" => new { pivotTableName = settings.PivotTableName, tableName = settings.TableName, destinationSheet = settings.DestSheet, destinationCell = settings.DestCell },
            "delete" => new { pivotTableName = settings.PivotTableName },
            "refresh" => new { pivotTableName = settings.PivotTableName },

            // PivotTableFieldAction
            "list-fields" => new { pivotTableName = settings.PivotTableName },
            "add-row-field" => new { pivotTableName = settings.PivotTableName, fieldName = settings.FieldName, position = settings.Position },
            "add-column-field" => new { pivotTableName = settings.PivotTableName, fieldName = settings.FieldName, position = settings.Position },
            "add-value-field" => new { pivotTableName = settings.PivotTableName, fieldName = settings.FieldName, aggregationFunction = settings.AggregationFunction, customName = settings.CustomName },
            "add-filter-field" => new { pivotTableName = settings.PivotTableName, fieldName = settings.FieldName, position = settings.Position },
            "remove-field" => new { pivotTableName = settings.PivotTableName, fieldName = settings.FieldName },
            "set-field-function" => new { pivotTableName = settings.PivotTableName, fieldName = settings.FieldName, aggregationFunction = settings.AggregationFunction },
            "set-field-name" => new { pivotTableName = settings.PivotTableName, fieldName = settings.FieldName, customName = settings.CustomName },
            "set-field-format" => new { pivotTableName = settings.PivotTableName, fieldName = settings.FieldName, numberFormat = settings.NumberFormat },
            "set-field-filter" => new { pivotTableName = settings.PivotTableName, fieldName = settings.FieldName, selectedValues = ParseStringList(settings.SelectedValues) },
            "sort-field" => new { pivotTableName = settings.PivotTableName, fieldName = settings.FieldName, ascending = settings.Ascending },
            "group-by-date" => new { pivotTableName = settings.PivotTableName, fieldName = settings.FieldName, interval = settings.Interval },
            "group-by-numeric" => new { pivotTableName = settings.PivotTableName, fieldName = settings.FieldName, start = settings.Start, end = settings.End, intervalSize = settings.IntervalSize },

            // PivotTableCalcAction
            "list-calculated-fields" => new { pivotTableName = settings.PivotTableName },
            "create-calculated-field" => new { pivotTableName = settings.PivotTableName, fieldName = settings.FieldName, formula = settings.Formula },
            "delete-calculated-field" => new { pivotTableName = settings.PivotTableName, fieldName = settings.FieldName },
            "list-calculated-members" => new { pivotTableName = settings.PivotTableName },
            "create-calculated-member" => new { pivotTableName = settings.PivotTableName, memberName = settings.MemberName, formula = settings.Formula, memberType = settings.MemberType, solveOrder = settings.SolveOrder, displayFolder = settings.DisplayFolder, numberFormat = settings.NumberFormat },
            "delete-calculated-member" => new { pivotTableName = settings.PivotTableName, memberName = settings.MemberName },
            "set-layout" => new { pivotTableName = settings.PivotTableName, layoutType = settings.LayoutType },
            "set-subtotals" => new { pivotTableName = settings.PivotTableName, fieldName = settings.FieldName, showSubtotals = settings.ShowSubtotals },
            "set-grand-totals" => new { pivotTableName = settings.PivotTableName, showRowGrandTotals = settings.ShowRowGrandTotals, showColumnGrandTotals = settings.ShowColumnGrandTotals },
            "get-data" => new { pivotTableName = settings.PivotTableName },

            _ => new { pivotTableName = settings.PivotTableName }
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

    private static List<string>? ParseStringList(string? input)
    {
        if (string.IsNullOrWhiteSpace(input)) return null;
        // Try JSON array first, fall back to comma-separated
        try
        {
            return JsonSerializer.Deserialize<List<string>>(input);
        }
        catch
        {
            return [.. input.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)];
        }
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<ACTION>")]
        [Description("The action to perform (e.g., list, create-from-range, refresh)")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        [Description("Session ID from 'session open' command")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--pivot-table <NAME>")]
        [Description("PivotTable name")]
        public string? PivotTableName { get; init; }

        [CommandOption("--table <NAME>")]
        [Description("Source table name for PivotTable creation")]
        public string? TableName { get; init; }

        [CommandOption("--source-sheet <NAME>")]
        [Description("Worksheet containing source data")]
        public string? SourceSheet { get; init; }

        [CommandOption("--source-range <ADDRESS>")]
        [Description("Source data range address")]
        public string? SourceRange { get; init; }

        [CommandOption("--dest-sheet <NAME>")]
        [Description("Destination worksheet for PivotTable")]
        public string? DestSheet { get; init; }

        [CommandOption("--dest-cell <ADDRESS>")]
        [Description("Destination cell for PivotTable placement")]
        public string? DestCell { get; init; }

        // PivotTableFieldAction settings
        [CommandOption("--field <NAME>")]
        [Description("Field name for add/remove/configure operations")]
        public string? FieldName { get; init; }

        [CommandOption("--position <NUMBER>")]
        [Description("Field position (0-based index)")]
        public int? Position { get; init; }

        [CommandOption("--function <NAME>")]
        [Description("Aggregation function (Sum, Count, Average, Max, Min, etc.)")]
        public string? AggregationFunction { get; init; }

        [CommandOption("--custom-name <NAME>")]
        [Description("Custom display name for field")]
        public string? CustomName { get; init; }

        [CommandOption("--number-format <FORMAT>")]
        [Description("Number format code (e.g., #,##0.00)")]
        public string? NumberFormat { get; init; }

        [CommandOption("--selected-values <VALUES>")]
        [Description("Comma-separated or JSON array of values for filter")]
        public string? SelectedValues { get; init; }

        [CommandOption("--ascending")]
        [Description("Sort ascending (default: true)")]
        public bool Ascending { get; init; } = true;

        [CommandOption("--interval <INTERVAL>")]
        [Description("Date grouping interval (Days, Months, Quarters, Years)")]
        public string? Interval { get; init; }

        [CommandOption("--start <NUMBER>")]
        [Description("Numeric grouping start value")]
        public double? Start { get; init; }

        [CommandOption("--end <NUMBER>")]
        [Description("Numeric grouping end value")]
        public double? End { get; init; }

        [CommandOption("--interval-size <NUMBER>")]
        [Description("Numeric grouping interval size")]
        public double? IntervalSize { get; init; }

        // PivotTableCalcAction settings
        [CommandOption("--formula <FORMULA>")]
        [Description("Formula for calculated field/member")]
        public string? Formula { get; init; }

        [CommandOption("--member-name <NAME>")]
        [Description("Calculated member name")]
        public string? MemberName { get; init; }

        [CommandOption("--member-type <TYPE>")]
        [Description("Calculated member type")]
        public string? MemberType { get; init; }

        [CommandOption("--solve-order <NUMBER>")]
        [Description("Calculated member solve order")]
        public int? SolveOrder { get; init; }

        [CommandOption("--display-folder <FOLDER>")]
        [Description("Calculated member display folder")]
        public string? DisplayFolder { get; init; }

        [CommandOption("--layout-type <TYPE>")]
        [Description("PivotTable layout type (0=Compact, 1=Outline, 2=Tabular)")]
        public int? LayoutType { get; init; }

        [CommandOption("--show-subtotals")]
        [Description("Show subtotals for field")]
        public bool ShowSubtotals { get; init; } = true;

        [CommandOption("--show-row-grand-totals")]
        [Description("Show row grand totals")]
        public bool ShowRowGrandTotals { get; init; } = true;

        [CommandOption("--show-column-grand-totals")]
        [Description("Show column grand totals")]
        public bool ShowColumnGrandTotals { get; init; } = true;

        [CommandOption("--layout-style <STYLE>")]
        [Description("PivotTable layout style")]
        public string? LayoutStyle { get; init; }
    }
}
