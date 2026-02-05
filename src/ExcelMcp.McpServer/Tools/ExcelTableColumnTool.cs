using System.ComponentModel;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel Table column, filter, and sort operations.
/// </summary>
[McpServerToolType]
public static partial class TableColumnTool
{
    /// <summary>
    /// Table column, filtering, and sorting operations for Excel Tables (ListObjects).
    ///
    /// FILTERING:
    /// - 'apply-filter': Simple criteria filter (e.g., ">100", "=Active", "&lt;>Closed")
    /// - 'apply-filter-values': Filter by exact values (provide JSON array of values to include)
    /// - 'clear-filters': Remove all active filters
    /// - 'get-filters': See current filter state
    ///
    /// SORTING:
    /// - 'sort': Single column sort (ascending/descending)
    /// - 'sort-multi': Multi-column sort (provide JSON array of {columnName, ascending} objects)
    ///
    /// COLUMN MANAGEMENT:
    /// - 'add-column'/'remove-column'/'rename-column': Modify table structure
    ///
    /// NUMBER FORMATS: Use US locale format codes (e.g., "#,##0.00", "0%", "yyyy-mm-dd")
    ///
    /// Related: excel_table (table lifecycle and data operations)
    /// </summary>
    /// <param name="action">The column/filter/sort operation to perform</param>
    /// <param name="path">Full path to Excel file (for reference/logging)</param>
    /// <param name="sessionId">Session ID from excel_file 'open'. Required for all actions.</param>
    /// <param name="tableName">Name of the table to operate on. Required for all actions.</param>
    /// <param name="columnName">Column name to operate on. Required for: apply-filter, apply-filter-values, add-column, remove-column, rename-column, sort, get/set-column-number-format</param>
    /// <param name="newColumnName">New name for the column. Required for: rename-column</param>
    /// <param name="criteria">Filter criteria string (e.g., ">100", "=Active"). Required for: apply-filter. Table region for: get-structured-reference (All, Data, Headers, Totals, ThisRow)</param>
    /// <param name="filterValuesJson">JSON array for filtering or sorting. For 'apply-filter-values': ["value1","value2"]. For 'sort-multi': [{"columnName":"Col1","ascending":true}]</param>
    /// <param name="formatCode">Number format code in US locale (e.g., "#,##0.00", "0%"). Required for: set-column-number-format</param>
    /// <param name="columnPosition">1-based column position for 'add-column' (optional, defaults to end of table)</param>
    /// <param name="ascending">Sort order for 'sort' action. true=ascending (A-Z, 0-9), false=descending. Default: true</param>
    [McpServerTool(Name = "excel_table_column", Title = "Excel Table Column Operations", Destructive = true)]
    [McpMeta("category", "data")]
    [McpMeta("requiresSession", true)]
    public static partial string TableColumn(
        TableColumnAction action,
        string path,
        string sessionId,
        [DefaultValue(null)] string? tableName,
        [DefaultValue(null)] string? columnName,
        [DefaultValue(null)] string? newColumnName,
        [DefaultValue(null)] string? criteria,
        [DefaultValue(null)] string? filterValuesJson,
        [DefaultValue(null)] string? formatCode,
        [DefaultValue(null)] string? columnPosition,
        [DefaultValue(true)] bool ascending)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_table_column",
            ServiceRegistry.TableColumn.ToActionString(action),
            path,
            () =>
            {
                return action switch
                {
                    TableColumnAction.ApplyFilter => ForwardApplyFilter(sessionId, tableName, columnName, criteria),
                    TableColumnAction.ApplyFilterValues => ForwardApplyFilterValues(sessionId, tableName, columnName, filterValuesJson),
                    TableColumnAction.ClearFilters => ForwardClearFilters(sessionId, tableName),
                    TableColumnAction.GetFilters => ForwardGetFilters(sessionId, tableName),
                    TableColumnAction.AddColumn => ForwardAddColumn(sessionId, tableName, columnName, columnPosition),
                    TableColumnAction.RemoveColumn => ForwardRemoveColumn(sessionId, tableName, columnName),
                    TableColumnAction.RenameColumn => ForwardRenameColumn(sessionId, tableName, columnName, newColumnName),
                    TableColumnAction.GetStructuredReference => ForwardGetStructuredReference(sessionId, tableName, criteria, columnName),
                    TableColumnAction.Sort => ForwardSort(sessionId, tableName, columnName, ascending),
                    TableColumnAction.SortMulti => ForwardSortMulti(sessionId, tableName, filterValuesJson),
                    TableColumnAction.GetColumnNumberFormat => ForwardGetColumnNumberFormat(sessionId, tableName, columnName),
                    TableColumnAction.SetColumnNumberFormat => ForwardSetColumnNumberFormat(sessionId, tableName, columnName, formatCode),
                    _ => throw new ArgumentException(
                        $"Unknown action: {action} ({ServiceRegistry.TableColumn.ToActionString(action)})", nameof(action))
                };
            });
    }

    private static string ForwardApplyFilter(string sessionId, string? tableName, string? columnName, string? criteria)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "apply-filter");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "apply-filter");
        if (string.IsNullOrWhiteSpace(criteria)) ExcelToolsBase.ThrowMissingParameter(nameof(criteria), "apply-filter");

        return ExcelToolsBase.ForwardToService("table.apply-filter", sessionId, new { tableName, columnName, criteria });
    }

    private static string ForwardApplyFilterValues(string sessionId, string? tableName, string? columnName, string? filterValuesJson)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "apply-filter-values");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "apply-filter-values");
        if (string.IsNullOrWhiteSpace(filterValuesJson)) ExcelToolsBase.ThrowMissingParameter(nameof(filterValuesJson), "apply-filter-values");

        return ExcelToolsBase.ForwardToService("table.apply-filter-values", sessionId, new { tableName, columnName, filterValues = filterValuesJson });
    }

    private static string ForwardClearFilters(string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "clear-filters");

        return ExcelToolsBase.ForwardToService("table.clear-filters", sessionId, new { tableName });
    }

    private static string ForwardGetFilters(string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "get-filters");

        return ExcelToolsBase.ForwardToService("table.get-filters", sessionId, new { tableName });
    }

    private static string ForwardAddColumn(string sessionId, string? tableName, string? columnName, string? columnPosition)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "add-column");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "add-column");

        int? position = null;
        if (!string.IsNullOrWhiteSpace(columnPosition) && int.TryParse(columnPosition, out var pos))
        {
            position = pos;
        }

        return ExcelToolsBase.ForwardToService("table.add-column", sessionId, new { tableName, columnName, position });
    }

    private static string ForwardRemoveColumn(string sessionId, string? tableName, string? columnName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "remove-column");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "remove-column");

        return ExcelToolsBase.ForwardToService("table.remove-column", sessionId, new { tableName, columnName });
    }

    private static string ForwardRenameColumn(string sessionId, string? tableName, string? columnName, string? newColumnName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "rename-column");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "rename-column");
        if (string.IsNullOrWhiteSpace(newColumnName)) ExcelToolsBase.ThrowMissingParameter(nameof(newColumnName), "rename-column");

        return ExcelToolsBase.ForwardToService("table.rename-column", sessionId, new { tableName, columnName, newName = newColumnName });
    }

    private static string ForwardGetStructuredReference(string sessionId, string? tableName, string? region, string? columnName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "get-structured-reference");

        return ExcelToolsBase.ForwardToService("table.get-structured-reference", sessionId, new { tableName, region, columnName });
    }

    private static string ForwardSort(string sessionId, string? tableName, string? columnName, bool ascending)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "sort");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "sort");

        return ExcelToolsBase.ForwardToService("table.sort", sessionId, new { tableName, columnName, ascending });
    }

    private static string ForwardSortMulti(string sessionId, string? tableName, string? sortColumns)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "sort-multi");
        if (string.IsNullOrWhiteSpace(sortColumns)) ExcelToolsBase.ThrowMissingParameter("filterValuesJson", "sort-multi");

        return ExcelToolsBase.ForwardToService("table.sort-multi", sessionId, new { tableName, sortColumns });
    }

    private static string ForwardGetColumnNumberFormat(string sessionId, string? tableName, string? columnName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "get-column-number-format");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "get-column-number-format");

        return ExcelToolsBase.ForwardToService("table.get-column-number-format", sessionId, new { tableName, columnName });
    }

    private static string ForwardSetColumnNumberFormat(string sessionId, string? tableName, string? columnName, string? formatCode)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "set-column-number-format");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "set-column-number-format");
        if (string.IsNullOrWhiteSpace(formatCode)) ExcelToolsBase.ThrowMissingParameter(nameof(formatCode), "set-column-number-format");

        return ExcelToolsBase.ForwardToService("table.set-column-number-format", sessionId, new { tableName, columnName, formatCode });
    }
}




