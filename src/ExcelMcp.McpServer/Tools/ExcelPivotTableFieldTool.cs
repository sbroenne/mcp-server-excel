using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for PivotTable field management operations
/// </summary>
[McpServerToolType]
public static partial class ExcelPivotTableFieldTool
{
    private static readonly JsonSerializerOptions JsonOptions = ExcelToolsBase.JsonOptions;

    /// <summary>
    /// PivotTable field management: add/remove/configure fields, filtering, sorting, and grouping.
    ///
    /// IMPORTANT: Field operations modify structure only. Call excel_pivottable(refresh) after
    /// configuring fields to update the visual display, especially for OLAP/Data Model PivotTables.
    ///
    /// FIELD AREAS:
    /// - Row fields: Group data by categories (add-row-field)
    /// - Column fields: Create column headers (add-column-field)
    /// - Value fields: Aggregate numeric data with Sum, Count, Average, etc. (add-value-field)
    /// - Filter fields: Add report-level filters (add-filter-field)
    ///
    /// AGGREGATION FUNCTIONS:
    /// Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP
    ///
    /// GROUPING:
    /// - Date fields: Group by Days, Months, Quarters, Years (group-by-date)
    /// - Numeric fields: Group by ranges with start/end/interval (group-by-numeric)
    ///
    /// NUMBER FORMAT: Use US format codes like '#,##0.00' for currency or '0.00%' for percentages.
    ///
    /// RELATED TOOLS:
    /// - excel_pivottable: Create/delete/refresh PivotTables
    /// - excel_pivottable_calc: Calculated fields, layout options, subtotals
    /// </summary>
    /// <param name="action">The field operation to perform</param>
    /// <param name="sessionId">Session identifier returned from excel_file open action</param>
    /// <param name="pivotTableName">Name of the PivotTable to modify</param>
    /// <param name="fieldName">Name of the field to add, remove, or configure</param>
    /// <param name="aggregationFunction">Aggregation function for value fields: Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP</param>
    /// <param name="customName">Custom display name for the field in the PivotTable</param>
    /// <param name="numberFormat">Number format code in US format, e.g., '#,##0.00' for currency, '0.00%' for percentage</param>
    /// <param name="position">1-based position for row/column field ordering</param>
    /// <param name="filterValues">JSON array of values to filter by, e.g., '["North","South"]' to show only those items</param>
    /// <param name="sortDirection">Sort direction for field items: Ascending or Descending</param>
    /// <param name="dateGroupingInterval">Date grouping interval: Days, Months, Quarters, or Years</param>
    /// <param name="numericGroupingStart">Starting value for numeric grouping ranges</param>
    /// <param name="numericGroupingEnd">Ending value for numeric grouping ranges</param>
    /// <param name="numericGroupingInterval">Interval size for numeric grouping (must be greater than 0)</param>
    [McpServerTool(Name = "excel_pivottable_field", Title = "Excel PivotTable Field Operations", Destructive = true)]
    [McpMeta("category", "analysis")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelPivotTableField(
        PivotTableFieldAction action,
        string sessionId,
        string pivotTableName,
        [DefaultValue(null)] string? fieldName,
        [DefaultValue(null)] string? aggregationFunction,
        [DefaultValue(null)] string? customName,
        [DefaultValue(null)] string? numberFormat,
        [DefaultValue(null)] int? position,
        [DefaultValue(null)] string? filterValues,
        [DefaultValue(null)] string? sortDirection,
        [DefaultValue(null)] string? dateGroupingInterval,
        [DefaultValue(null)] double? numericGroupingStart,
        [DefaultValue(null)] double? numericGroupingEnd,
        [DefaultValue(null)] double? numericGroupingInterval)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_pivottable_field",
            action.ToActionString(),
            () =>
            {
                var commands = new PivotTableCommands();

                return action switch
                {
                    PivotTableFieldAction.ListFields => ListFields(commands, sessionId, pivotTableName),
                    PivotTableFieldAction.AddRowField => AddRowField(commands, sessionId, pivotTableName, fieldName, position),
                    PivotTableFieldAction.AddColumnField => AddColumnField(commands, sessionId, pivotTableName, fieldName, position),
                    PivotTableFieldAction.AddValueField => AddValueField(commands, sessionId, pivotTableName, fieldName, aggregationFunction, customName),
                    PivotTableFieldAction.AddFilterField => AddFilterField(commands, sessionId, pivotTableName, fieldName),
                    PivotTableFieldAction.RemoveField => RemoveField(commands, sessionId, pivotTableName, fieldName),
                    PivotTableFieldAction.SetFieldFunction => SetFieldFunction(commands, sessionId, pivotTableName, fieldName, aggregationFunction),
                    PivotTableFieldAction.SetFieldName => SetFieldName(commands, sessionId, pivotTableName, fieldName, customName),
                    PivotTableFieldAction.SetFieldFormat => SetFieldFormat(commands, sessionId, pivotTableName, fieldName, numberFormat),
                    PivotTableFieldAction.SetFieldFilter => SetFieldFilter(commands, sessionId, pivotTableName, fieldName, filterValues),
                    PivotTableFieldAction.SortField => SortField(commands, sessionId, pivotTableName, fieldName, sortDirection),
                    PivotTableFieldAction.GroupByDate => GroupByDate(commands, sessionId, pivotTableName, fieldName, dateGroupingInterval),
                    PivotTableFieldAction.GroupByNumeric => GroupByNumeric(commands, sessionId, pivotTableName, fieldName, numericGroupingStart, numericGroupingEnd, numericGroupingInterval),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string ListFields(PivotTableCommands commands, string sessionId, string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "list-fields");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.ListFields(batch, pivotTableName!));

        return JsonSerializer.Serialize(new { result.Success, result.Fields, result.ErrorMessage }, JsonOptions);
    }

    private static string AddRowField(PivotTableCommands commands, string sessionId, string? pivotTableName, string? fieldName, int? position)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "add-row-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fieldName", "add-row-field");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.AddRowField(batch, pivotTableName!, fieldName!, position));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.Position,
            result.Function,
            result.NumberFormat,
            result.AvailableValues,
            result.SampleValue,
            result.DataType,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string AddColumnField(PivotTableCommands commands, string sessionId, string? pivotTableName, string? fieldName, int? position)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "add-column-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fieldName", "add-column-field");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.AddColumnField(batch, pivotTableName!, fieldName!, position));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.Position,
            result.Function,
            result.NumberFormat,
            result.AvailableValues,
            result.SampleValue,
            result.DataType,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string AddValueField(PivotTableCommands commands, string sessionId, string? pivotTableName, string? fieldName, string? aggregationFunction, string? customName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "add-value-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fieldName", "add-value-field");

        AggregationFunction function = AggregationFunction.Sum;
        if (!string.IsNullOrEmpty(aggregationFunction) &&
            !Enum.TryParse(aggregationFunction, true, out function))
        {
            throw new ArgumentException($"Invalid aggregationFunction '{aggregationFunction}'. Valid: Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP", nameof(aggregationFunction));
        }

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.AddValueField(batch, pivotTableName!, fieldName!, function, customName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.Position,
            result.Function,
            result.NumberFormat,
            result.AvailableValues,
            result.SampleValue,
            result.DataType,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string AddFilterField(PivotTableCommands commands, string sessionId, string? pivotTableName, string? fieldName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "add-filter-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fieldName", "add-filter-field");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.AddFilterField(batch, pivotTableName!, fieldName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.Position,
            result.Function,
            result.NumberFormat,
            result.AvailableValues,
            result.SampleValue,
            result.DataType,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string RemoveField(PivotTableCommands commands, string sessionId, string? pivotTableName, string? fieldName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "remove-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fieldName", "remove-field");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.RemoveField(batch, pivotTableName!, fieldName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.Position,
            result.Function,
            result.NumberFormat,
            result.AvailableValues,
            result.SampleValue,
            result.DataType,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string SetFieldFunction(PivotTableCommands commands, string sessionId, string? pivotTableName, string? fieldName, string? aggregationFunction)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "set-field-function");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fieldName", "set-field-function");
        if (string.IsNullOrWhiteSpace(aggregationFunction))
            ExcelToolsBase.ThrowMissingParameter("aggregationFunction", "set-field-function");

        if (!Enum.TryParse<AggregationFunction>(aggregationFunction!, true, out var function))
        {
            throw new ArgumentException($"Invalid aggregationFunction '{aggregationFunction}'. Valid: Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP", nameof(aggregationFunction));
        }

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetFieldFunction(batch, pivotTableName!, fieldName!, function));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.Position,
            result.Function,
            result.NumberFormat,
            result.AvailableValues,
            result.SampleValue,
            result.DataType,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string SetFieldName(PivotTableCommands commands, string sessionId, string? pivotTableName, string? fieldName, string? customName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "set-field-name");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fieldName", "set-field-name");
        if (string.IsNullOrWhiteSpace(customName))
            ExcelToolsBase.ThrowMissingParameter("customName", "set-field-name");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetFieldName(batch, pivotTableName!, fieldName!, customName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.Position,
            result.Function,
            result.NumberFormat,
            result.AvailableValues,
            result.SampleValue,
            result.DataType,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string SetFieldFormat(PivotTableCommands commands, string sessionId, string? pivotTableName, string? fieldName, string? numberFormat)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "set-field-format");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fieldName", "set-field-format");
        if (string.IsNullOrWhiteSpace(numberFormat))
            ExcelToolsBase.ThrowMissingParameter("numberFormat", "set-field-format");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetFieldFormat(batch, pivotTableName!, fieldName!, numberFormat!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.Position,
            result.Function,
            result.NumberFormat,
            result.AvailableValues,
            result.SampleValue,
            result.DataType,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string SetFieldFilter(PivotTableCommands commands, string sessionId, string? pivotTableName, string? fieldName, string? filterValues)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "set-field-filter");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fieldName", "set-field-filter");
        if (string.IsNullOrWhiteSpace(filterValues))
            ExcelToolsBase.ThrowMissingParameter("filterValues", "set-field-filter");

        List<string> values;
        try
        {
            values = JsonSerializer.Deserialize<List<string>>(filterValues!) ?? [];
        }
        catch (JsonException ex)
        {
            throw new ArgumentException($"Invalid filterValues JSON: {ex.Message}. Expected: '[\"value1\",\"value2\"]'", nameof(filterValues));
        }

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetFieldFilter(batch, pivotTableName!, fieldName!, values));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.SelectedItems,
            result.AvailableItems,
            result.VisibleRowCount,
            result.TotalRowCount,
            result.ShowAll,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string SortField(PivotTableCommands commands, string sessionId, string? pivotTableName, string? fieldName, string? sortDirection)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "sort-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fieldName", "sort-field");

        SortDirection direction = SortDirection.Ascending;
        if (!string.IsNullOrEmpty(sortDirection) &&
            !Enum.TryParse(sortDirection, true, out direction))
        {
            throw new ArgumentException($"Invalid sortDirection '{sortDirection}'. Valid: Ascending, Descending", nameof(sortDirection));
        }

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SortField(batch, pivotTableName!, fieldName!, direction));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.Position,
            result.Function,
            result.NumberFormat,
            result.AvailableValues,
            result.SampleValue,
            result.DataType,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string GroupByDate(PivotTableCommands commands, string sessionId, string? pivotTableName, string? fieldName, string? dateGroupingInterval)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for group-by-date action", nameof(pivotTableName));
        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fieldName is required for group-by-date action", nameof(fieldName));
        if (string.IsNullOrEmpty(dateGroupingInterval))
            throw new ArgumentException("dateGroupingInterval is required for group-by-date action", nameof(dateGroupingInterval));

        if (!Enum.TryParse<DateGroupingInterval>(dateGroupingInterval, true, out var interval))
        {
            throw new ArgumentException($"Invalid dateGroupingInterval '{dateGroupingInterval}'. Valid: Days, Months, Quarters, Years", nameof(dateGroupingInterval));
        }

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.GroupByDate(batch, pivotTableName!, fieldName!, interval));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.ErrorMessage,
            result.WorkflowHint
        }, JsonOptions);
    }

    private static string GroupByNumeric(PivotTableCommands commands, string sessionId, string? pivotTableName, string? fieldName, double? numericGroupingStart, double? numericGroupingEnd, double? numericGroupingInterval)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("pivotTableName is required for group-by-numeric action", nameof(pivotTableName));
        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fieldName is required for group-by-numeric action", nameof(fieldName));
        if (!numericGroupingInterval.HasValue || numericGroupingInterval.Value <= 0)
            throw new ArgumentException("numericGroupingInterval is required and must be > 0 for group-by-numeric action", nameof(numericGroupingInterval));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.GroupByNumeric(batch, pivotTableName!, fieldName!, numericGroupingStart, numericGroupingEnd, numericGroupingInterval.Value));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FieldName,
            result.CustomName,
            result.Area,
            result.ErrorMessage,
            result.WorkflowHint
        }, JsonOptions);
    }
}
