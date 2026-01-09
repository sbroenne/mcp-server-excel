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
    /// PivotTable field management - add/remove/configure fields, filtering, sorting, grouping.
    /// Related: excel_pivottable (lifecycle), excel_pivottable_calc (calculated/layout)
    /// </summary>
    /// <param name="action">Action</param>
    /// <param name="sid">Session ID</param>
    /// <param name="ptn">PivotTable name</param>
    /// <param name="fn">Field name</param>
    /// <param name="agg">Aggregation: Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP</param>
    /// <param name="cn">Custom display name</param>
    /// <param name="nf">Number format (US format: '#,##0.00', '0.00%')</param>
    /// <param name="pos">Position (1-based)</param>
    /// <param name="fv">Filter values JSON array: '["val1","val2"]'</param>
    /// <param name="sd">Sort direction: Ascending, Descending</param>
    /// <param name="dgi">Date grouping: Days, Months, Quarters, Years</param>
    /// <param name="ngs">Numeric grouping start</param>
    /// <param name="nge">Numeric grouping end</param>
    /// <param name="ngi">Numeric grouping interval</param>
    [McpServerTool(Name = "excel_pivottable_field", Title = "Excel PivotTable Field Operations")]
    [McpMeta("category", "analysis")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelPivotTableField(
        PivotTableFieldAction action,
        string sid,
        string ptn,
        [DefaultValue(null)] string? fn,
        [DefaultValue(null)] string? agg,
        [DefaultValue(null)] string? cn,
        [DefaultValue(null)] string? nf,
        [DefaultValue(null)] int? pos,
        [DefaultValue(null)] string? fv,
        [DefaultValue(null)] string? sd,
        [DefaultValue(null)] string? dgi,
        [DefaultValue(null)] double? ngs,
        [DefaultValue(null)] double? nge,
        [DefaultValue(null)] double? ngi)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_pivottable_field",
            action.ToActionString(),
            () =>
            {
                var commands = new PivotTableCommands();

                return action switch
                {
                    PivotTableFieldAction.ListFields => ListFields(commands, sid, ptn),
                    PivotTableFieldAction.AddRowField => AddRowField(commands, sid, ptn, fn, pos),
                    PivotTableFieldAction.AddColumnField => AddColumnField(commands, sid, ptn, fn, pos),
                    PivotTableFieldAction.AddValueField => AddValueField(commands, sid, ptn, fn, agg, cn),
                    PivotTableFieldAction.AddFilterField => AddFilterField(commands, sid, ptn, fn),
                    PivotTableFieldAction.RemoveField => RemoveField(commands, sid, ptn, fn),
                    PivotTableFieldAction.SetFieldFunction => SetFieldFunction(commands, sid, ptn, fn, agg),
                    PivotTableFieldAction.SetFieldName => SetFieldName(commands, sid, ptn, fn, cn),
                    PivotTableFieldAction.SetFieldFormat => SetFieldFormat(commands, sid, ptn, fn, nf),
                    PivotTableFieldAction.SetFieldFilter => SetFieldFilter(commands, sid, ptn, fn, fv),
                    PivotTableFieldAction.SortField => SortField(commands, sid, ptn, fn, sd),
                    PivotTableFieldAction.GroupByDate => GroupByDate(commands, sid, ptn, fn, dgi),
                    PivotTableFieldAction.GroupByNumeric => GroupByNumeric(commands, sid, ptn, fn, ngs, nge, ngi),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string ListFields(PivotTableCommands commands, string sessionId, string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("ptn", "list-fields");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.ListFields(batch, pivotTableName!));

        return JsonSerializer.Serialize(new { result.Success, result.Fields, result.ErrorMessage }, JsonOptions);
    }

    private static string AddRowField(PivotTableCommands commands, string sessionId, string? pivotTableName, string? fieldName, int? position)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("ptn", "add-row-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fn", "add-row-field");

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
            ExcelToolsBase.ThrowMissingParameter("ptn", "add-column-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fn", "add-column-field");

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
            ExcelToolsBase.ThrowMissingParameter("ptn", "add-value-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fn", "add-value-field");

        AggregationFunction function = AggregationFunction.Sum;
        if (!string.IsNullOrEmpty(aggregationFunction) &&
            !Enum.TryParse(aggregationFunction, true, out function))
        {
            throw new ArgumentException($"Invalid aggregation '{aggregationFunction}'. Valid: Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP", nameof(aggregationFunction));
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
            ExcelToolsBase.ThrowMissingParameter("ptn", "add-filter-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fn", "add-filter-field");

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
            ExcelToolsBase.ThrowMissingParameter("ptn", "remove-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fn", "remove-field");

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
            ExcelToolsBase.ThrowMissingParameter("ptn", "set-field-function");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fn", "set-field-function");
        if (string.IsNullOrWhiteSpace(aggregationFunction))
            ExcelToolsBase.ThrowMissingParameter("agg", "set-field-function");

        if (!Enum.TryParse<AggregationFunction>(aggregationFunction!, true, out var function))
        {
            throw new ArgumentException($"Invalid aggregation '{aggregationFunction}'. Valid: Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP", nameof(aggregationFunction));
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
            ExcelToolsBase.ThrowMissingParameter("ptn", "set-field-name");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fn", "set-field-name");
        if (string.IsNullOrWhiteSpace(customName))
            ExcelToolsBase.ThrowMissingParameter("cn", "set-field-name");

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
            ExcelToolsBase.ThrowMissingParameter("ptn", "set-field-format");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fn", "set-field-format");
        if (string.IsNullOrWhiteSpace(numberFormat))
            ExcelToolsBase.ThrowMissingParameter("nf", "set-field-format");

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
            ExcelToolsBase.ThrowMissingParameter("ptn", "set-field-filter");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fn", "set-field-filter");
        if (string.IsNullOrWhiteSpace(filterValues))
            ExcelToolsBase.ThrowMissingParameter("fv", "set-field-filter");

        List<string> values;
        try
        {
            values = JsonSerializer.Deserialize<List<string>>(filterValues!) ?? [];
        }
        catch (JsonException ex)
        {
            throw new ArgumentException($"Invalid fv JSON: {ex.Message}. Expected: '[\"value1\",\"value2\"]'", nameof(filterValues));
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
            ExcelToolsBase.ThrowMissingParameter("ptn", "sort-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fn", "sort-field");

        SortDirection direction = SortDirection.Ascending;
        if (!string.IsNullOrEmpty(sortDirection) &&
            !Enum.TryParse(sortDirection, true, out direction))
        {
            throw new ArgumentException($"Invalid sort direction '{sortDirection}'. Valid: Ascending, Descending", nameof(sortDirection));
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
            throw new ArgumentException("ptn is required for group-by-date action", nameof(pivotTableName));
        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fn is required for group-by-date action", nameof(fieldName));
        if (string.IsNullOrEmpty(dateGroupingInterval))
            throw new ArgumentException("dgi is required for group-by-date action", nameof(dateGroupingInterval));

        if (!Enum.TryParse<DateGroupingInterval>(dateGroupingInterval, true, out var interval))
        {
            throw new ArgumentException($"Invalid date grouping '{dateGroupingInterval}'. Valid: Days, Months, Quarters, Years", nameof(dateGroupingInterval));
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

    private static string GroupByNumeric(PivotTableCommands commands, string sessionId, string? pivotTableName, string? fieldName, double? start, double? end, double? interval)
    {
        if (string.IsNullOrEmpty(pivotTableName))
            throw new ArgumentException("ptn is required for group-by-numeric action", nameof(pivotTableName));
        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fn is required for group-by-numeric action", nameof(fieldName));
        if (!interval.HasValue || interval.Value <= 0)
            throw new ArgumentException("ngi is required and must be > 0 for group-by-numeric action", nameof(interval));

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.GroupByNumeric(batch, pivotTableName!, fieldName!, start, end, interval.Value));

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
