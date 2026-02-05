using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for PivotTable field management operations
/// </summary>
[McpServerToolType]
public static partial class ExcelPivotTableFieldTool
{
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
            ServiceRegistry.PivotTableField.ToActionString(action),
            () => action switch
            {
                PivotTableFieldAction.ListFields => ForwardListFields(sessionId, pivotTableName),
                PivotTableFieldAction.AddRowField => ForwardAddRowField(sessionId, pivotTableName, fieldName, position),
                PivotTableFieldAction.AddColumnField => ForwardAddColumnField(sessionId, pivotTableName, fieldName, position),
                PivotTableFieldAction.AddValueField => ForwardAddValueField(sessionId, pivotTableName, fieldName, aggregationFunction, customName),
                PivotTableFieldAction.AddFilterField => ForwardAddFilterField(sessionId, pivotTableName, fieldName),
                PivotTableFieldAction.RemoveField => ForwardRemoveField(sessionId, pivotTableName, fieldName),
                PivotTableFieldAction.SetFieldFunction => ForwardSetFieldFunction(sessionId, pivotTableName, fieldName, aggregationFunction),
                PivotTableFieldAction.SetFieldName => ForwardSetFieldName(sessionId, pivotTableName, fieldName, customName),
                PivotTableFieldAction.SetFieldFormat => ForwardSetFieldFormat(sessionId, pivotTableName, fieldName, numberFormat),
                PivotTableFieldAction.SetFieldFilter => ForwardSetFieldFilter(sessionId, pivotTableName, fieldName, filterValues),
                PivotTableFieldAction.SortField => ForwardSortField(sessionId, pivotTableName, fieldName, sortDirection),
                PivotTableFieldAction.GroupByDate => ForwardGroupByDate(sessionId, pivotTableName, fieldName, dateGroupingInterval),
                PivotTableFieldAction.GroupByNumeric => ForwardGroupByNumeric(sessionId, pivotTableName, fieldName, numericGroupingStart, numericGroupingEnd, numericGroupingInterval),
                _ => throw new ArgumentException($"Unknown action: {action} ({ServiceRegistry.PivotTableField.ToActionString(action)})", nameof(action))
            });
    }

    // === SERVICE FORWARDING METHODS ===

    private static string ForwardListFields(string sessionId, string? pivotTableName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "list-fields");

        return ExcelToolsBase.ForwardToService("pivottablefield.list-fields", sessionId, new { pivotTableName });
    }

    private static string ForwardAddRowField(string sessionId, string? pivotTableName, string? fieldName, int? position)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "add-row-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fieldName", "add-row-field");

        return ExcelToolsBase.ForwardToService("pivottablefield.add-row-field", sessionId, new { pivotTableName, fieldName, position });
    }

    private static string ForwardAddColumnField(string sessionId, string? pivotTableName, string? fieldName, int? position)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "add-column-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fieldName", "add-column-field");

        return ExcelToolsBase.ForwardToService("pivottablefield.add-column-field", sessionId, new { pivotTableName, fieldName, position });
    }

    private static string ForwardAddValueField(string sessionId, string? pivotTableName, string? fieldName, string? aggregationFunction, string? customName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "add-value-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fieldName", "add-value-field");

        return ExcelToolsBase.ForwardToService("pivottablefield.add-value-field", sessionId, new
        {
            pivotTableName,
            fieldName,
            aggregationFunction,
            customName
        });
    }

    private static string ForwardAddFilterField(string sessionId, string? pivotTableName, string? fieldName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "add-filter-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fieldName", "add-filter-field");

        return ExcelToolsBase.ForwardToService("pivottablefield.add-filter-field", sessionId, new { pivotTableName, fieldName });
    }

    private static string ForwardRemoveField(string sessionId, string? pivotTableName, string? fieldName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "remove-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fieldName", "remove-field");

        return ExcelToolsBase.ForwardToService("pivottablefield.remove-field", sessionId, new { pivotTableName, fieldName });
    }

    private static string ForwardSetFieldFunction(string sessionId, string? pivotTableName, string? fieldName, string? aggregationFunction)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "set-field-function");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fieldName", "set-field-function");
        if (string.IsNullOrWhiteSpace(aggregationFunction))
            ExcelToolsBase.ThrowMissingParameter("aggregationFunction", "set-field-function");

        return ExcelToolsBase.ForwardToService("pivottablefield.set-field-function", sessionId, new
        {
            pivotTableName,
            fieldName,
            aggregationFunction
        });
    }

    private static string ForwardSetFieldName(string sessionId, string? pivotTableName, string? fieldName, string? customName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "set-field-name");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fieldName", "set-field-name");
        if (string.IsNullOrWhiteSpace(customName))
            ExcelToolsBase.ThrowMissingParameter("customName", "set-field-name");

        return ExcelToolsBase.ForwardToService("pivottablefield.set-field-name", sessionId, new
        {
            pivotTableName,
            fieldName,
            customName
        });
    }

    private static string ForwardSetFieldFormat(string sessionId, string? pivotTableName, string? fieldName, string? numberFormat)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "set-field-format");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fieldName", "set-field-format");
        if (string.IsNullOrWhiteSpace(numberFormat))
            ExcelToolsBase.ThrowMissingParameter("numberFormat", "set-field-format");

        return ExcelToolsBase.ForwardToService("pivottablefield.set-field-format", sessionId, new
        {
            pivotTableName,
            fieldName,
            numberFormat
        });
    }

    private static string ForwardSetFieldFilter(string sessionId, string? pivotTableName, string? fieldName, string? filterValues)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "set-field-filter");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fieldName", "set-field-filter");
        if (string.IsNullOrWhiteSpace(filterValues))
            ExcelToolsBase.ThrowMissingParameter("filterValues", "set-field-filter");

        List<string> selectedValues;
        try
        {
            selectedValues = JsonSerializer.Deserialize<List<string>>(filterValues!) ?? [];
        }
        catch (JsonException ex)
        {
            throw new ArgumentException($"Invalid filterValues JSON: {ex.Message}. Expected: '[\"value1\",\"value2\"]'", nameof(filterValues));
        }

        return ExcelToolsBase.ForwardToService("pivottablefield.set-field-filter", sessionId, new
        {
            pivotTableName,
            fieldName,
            selectedValues
        });
    }

    private static string ForwardSortField(string sessionId, string? pivotTableName, string? fieldName, string? sortDirection)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "sort-field");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fieldName", "sort-field");

        // Parse sortDirection to boolean for the service
        // Service expects "ascending" boolean, not "sortDirection" string
        bool ascending = true;
        if (!string.IsNullOrEmpty(sortDirection))
        {
            ascending = sortDirection.Equals("Ascending", StringComparison.OrdinalIgnoreCase);
        }

        return ExcelToolsBase.ForwardToService("pivottablefield.sort-field", sessionId, new
        {
            pivotTableName,
            fieldName,
            ascending
        });
    }

    private static string ForwardGroupByDate(string sessionId, string? pivotTableName, string? fieldName, string? dateGroupingInterval)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "group-by-date");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fieldName", "group-by-date");
        if (string.IsNullOrWhiteSpace(dateGroupingInterval))
            ExcelToolsBase.ThrowMissingParameter("dateGroupingInterval", "group-by-date");

        return ExcelToolsBase.ForwardToService("pivottablefield.group-by-date", sessionId, new
        {
            pivotTableName,
            fieldName,
            interval = dateGroupingInterval
        });
    }

    private static string ForwardGroupByNumeric(string sessionId, string? pivotTableName, string? fieldName, double? numericGroupingStart, double? numericGroupingEnd, double? numericGroupingInterval)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter("pivotTableName", "group-by-numeric");
        if (string.IsNullOrWhiteSpace(fieldName))
            ExcelToolsBase.ThrowMissingParameter("fieldName", "group-by-numeric");
        if (!numericGroupingInterval.HasValue || numericGroupingInterval.Value <= 0)
            throw new ArgumentException("numericGroupingInterval is required and must be > 0 for group-by-numeric action", nameof(numericGroupingInterval));

        return ExcelToolsBase.ForwardToService("pivottablefield.group-by-numeric", sessionId, new
        {
            pivotTableName,
            fieldName,
            start = numericGroupingStart,
            end = numericGroupingEnd,
            intervalSize = numericGroupingInterval
        });
    }
}




