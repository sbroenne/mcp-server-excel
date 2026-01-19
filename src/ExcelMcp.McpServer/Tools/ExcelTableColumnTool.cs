using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;

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
    /// <param name="excelPath">Full path to Excel file (for reference/logging)</param>
    /// <param name="sessionId">Session ID from excel_file 'open'. Required for all actions.</param>
    /// <param name="tableName">Name of the table to operate on. Required for all actions.</param>
    /// <param name="columnName">Column name to operate on. Required for: apply-filter, apply-filter-values, add-column, remove-column, rename-column, sort, get/set-column-number-format</param>
    /// <param name="newColumnName">New name for the column. Required for: rename-column</param>
    /// <param name="criteria">Filter criteria string (e.g., ">100", "=Active"). Required for: apply-filter. Table region for: get-structured-reference (All, Data, Headers, Totals, ThisRow)</param>
    /// <param name="filterValuesJson">JSON array for filtering or sorting. For 'apply-filter-values': ["value1","value2"]. For 'sort-multi': [{"columnName":"Col1","ascending":true}]</param>
    /// <param name="formatCode">Number format code in US locale (e.g., "#,##0.00", "0%"). Required for: set-column-number-format</param>
    /// <param name="columnPosition">1-based column position for 'add-column' (optional, defaults to end of table)</param>
    /// <param name="ascending">Sort order for 'sort' action. true=ascending (A-Z, 0-9), false=descending. Default: true</param>
    [McpServerTool(Name = "excel_table_column", Title = "Excel Table Column Operations")]
    [McpMeta("category", "data")]
    [McpMeta("requiresSession", true)]
    public static partial string TableColumn(
        TableColumnAction action,
        string excelPath,
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
            action.ToActionString(),
            excelPath,
            () =>
            {
                var tableCommands = new TableCommands();

                return action switch
                {
                    TableColumnAction.ApplyFilter => ApplyFilter(tableCommands, sessionId, tableName, columnName, criteria),
                    TableColumnAction.ApplyFilterValues => ApplyFilterValues(tableCommands, sessionId, tableName, columnName, filterValuesJson),
                    TableColumnAction.ClearFilters => ClearFilters(tableCommands, sessionId, tableName),
                    TableColumnAction.GetFilters => GetFilters(tableCommands, sessionId, tableName),
                    TableColumnAction.AddColumn => AddColumn(tableCommands, sessionId, tableName, columnName, columnPosition),
                    TableColumnAction.RemoveColumn => RemoveColumn(tableCommands, sessionId, tableName, columnName),
                    TableColumnAction.RenameColumn => RenameColumn(tableCommands, sessionId, tableName, columnName, newColumnName),
                    TableColumnAction.GetStructuredReference => GetStructuredReference(tableCommands, sessionId, tableName, criteria, columnName),
                    TableColumnAction.Sort => SortTable(tableCommands, sessionId, tableName, columnName, ascending),
                    TableColumnAction.SortMulti => SortTableMulti(tableCommands, sessionId, tableName, filterValuesJson),
                    TableColumnAction.GetColumnNumberFormat => GetColumnNumberFormat(tableCommands, sessionId, tableName, columnName),
                    TableColumnAction.SetColumnNumberFormat => SetColumnNumberFormat(tableCommands, sessionId, tableName, columnName, formatCode),
                    _ => throw new ArgumentException(
                        $"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string ApplyFilter(TableCommands commands, string sessionId, string? tableName, string? columnName, string? criteria)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "apply-filter");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "apply-filter");
        if (string.IsNullOrWhiteSpace(criteria)) ExcelToolsBase.ThrowMissingParameter(nameof(criteria), "apply-filter");

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.ApplyFilter(batch, tableName!, columnName!, criteria!);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Filter applied successfully." }, ExcelToolsBase.JsonOptions);
        }
#pragma warning disable CA1031 // MCP protocol requires JSON error responses, not thrown exceptions
        catch (Exception ex)
#pragma warning restore CA1031
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string ApplyFilterValues(TableCommands commands, string sessionId, string? tableName, string? columnName, string? filterValuesJson)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "apply-filter-values");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "apply-filter-values");
        if (string.IsNullOrWhiteSpace(filterValuesJson)) ExcelToolsBase.ThrowMissingParameter(nameof(filterValuesJson), "apply-filter-values");

        List<string> filterValues;
        try
        {
            filterValues = JsonSerializer.Deserialize<List<string>>(filterValuesJson!) ?? [];
        }
        catch (JsonException ex)
        {
            throw new ArgumentException($"Invalid JSON array for filterValuesJson: {ex.Message}", nameof(filterValuesJson));
        }

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.ApplyFilter(batch, tableName!, columnName!, filterValues);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Filter applied successfully." }, ExcelToolsBase.JsonOptions);
        }
#pragma warning disable CA1031 // MCP protocol requires JSON error responses, not thrown exceptions
        catch (Exception ex)
#pragma warning restore CA1031
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string ClearFilters(TableCommands commands, string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "clear-filters");

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.ClearFilters(batch, tableName!);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Filters cleared successfully." }, ExcelToolsBase.JsonOptions);
        }
#pragma warning disable CA1031 // MCP protocol requires JSON error responses, not thrown exceptions
        catch (Exception ex)
#pragma warning restore CA1031
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string GetFilters(TableCommands commands, string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "get-filters");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetFilters(batch, tableName!));

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string AddColumn(TableCommands commands, string sessionId, string? tableName, string? columnName, string? columnPosition)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "add-column");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "add-column");

        int? position = null;
        if (!string.IsNullOrWhiteSpace(columnPosition))
        {
            if (int.TryParse(columnPosition, out int posVal))
            {
                position = posVal;
            }
            else
            {
                throw new ArgumentException($"Invalid columnPosition value: '{columnPosition}'. Must be a number.", nameof(columnPosition));
            }
        }

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.AddColumn(batch, tableName!, columnName!, position);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Column added successfully." }, ExcelToolsBase.JsonOptions);
        }
#pragma warning disable CA1031 // MCP protocol requires JSON error responses, not thrown exceptions
        catch (Exception ex)
#pragma warning restore CA1031
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string RemoveColumn(TableCommands commands, string sessionId, string? tableName, string? columnName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "remove-column");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "remove-column");

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.RemoveColumn(batch, tableName!, columnName!);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Column removed successfully." }, ExcelToolsBase.JsonOptions);
        }
#pragma warning disable CA1031 // MCP protocol requires JSON error responses, not thrown exceptions
        catch (Exception ex)
#pragma warning restore CA1031
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string RenameColumn(TableCommands commands, string sessionId, string? tableName, string? columnName, string? newColumnName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "rename-column");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "rename-column");
        if (string.IsNullOrWhiteSpace(newColumnName)) ExcelToolsBase.ThrowMissingParameter(nameof(newColumnName), "rename-column");

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.RenameColumn(batch, tableName!, columnName!, newColumnName!);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Column renamed successfully." }, ExcelToolsBase.JsonOptions);
        }
#pragma warning disable CA1031 // MCP protocol requires JSON error responses, not thrown exceptions
        catch (Exception ex)
#pragma warning restore CA1031
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string GetStructuredReference(TableCommands commands, string sessionId, string? tableName, string? region, string? columnName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "get-structured-reference");

        var tableRegion = TableRegion.Data;
        if (!string.IsNullOrWhiteSpace(region) && !Enum.TryParse(region, true, out tableRegion))
        {
            throw new ArgumentException($"Invalid region '{region}'. Valid values: All, Data, Headers, Totals, ThisRow", nameof(region));
        }

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetStructuredReference(batch, tableName!, tableRegion, columnName));

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SortTable(TableCommands commands, string sessionId, string? tableName, string? columnName, bool ascending)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "sort");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "sort");

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.Sort(batch, tableName!, columnName!, ascending);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Table sorted successfully." }, ExcelToolsBase.JsonOptions);
        }
#pragma warning disable CA1031 // MCP protocol requires JSON error responses, not thrown exceptions
        catch (Exception ex)
#pragma warning restore CA1031
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string SortTableMulti(TableCommands commands, string sessionId, string? tableName, string? sortColumnsJson)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "sort-multi");
        if (string.IsNullOrWhiteSpace(sortColumnsJson)) ExcelToolsBase.ThrowMissingParameter(nameof(sortColumnsJson), "sort-multi");

        List<TableSortColumn>? sortColumns;
        try
        {
            sortColumns = JsonSerializer.Deserialize<List<TableSortColumn>>(sortColumnsJson!);
            if (sortColumns == null || sortColumns.Count == 0)
            {
                throw new ArgumentException("sortColumnsJson must be a non-empty array", nameof(sortColumnsJson));
            }
        }
        catch (JsonException ex)
        {
            throw new ArgumentException($"Invalid sortColumnsJson: {ex.Message}", nameof(sortColumnsJson));
        }

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.Sort(batch, tableName!, sortColumns);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Table sorted successfully." }, ExcelToolsBase.JsonOptions);
        }
#pragma warning disable CA1031 // MCP protocol requires JSON error responses, not thrown exceptions
        catch (Exception ex)
#pragma warning restore CA1031
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string GetColumnNumberFormat(TableCommands commands, string sessionId, string? tableName, string? columnName)
    {
        if (string.IsNullOrEmpty(tableName))
            ExcelToolsBase.ThrowMissingParameter("tableName", "get-column-number-format");
        if (string.IsNullOrEmpty(columnName))
            ExcelToolsBase.ThrowMissingParameter("columnName", "get-column-number-format");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetColumnNumberFormat(batch, tableName!, columnName!));

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetColumnNumberFormat(TableCommands commands, string sessionId, string? tableName, string? columnName, string? formatCode)
    {
        if (string.IsNullOrEmpty(tableName))
            ExcelToolsBase.ThrowMissingParameter("tableName", "set-column-number-format");
        if (string.IsNullOrEmpty(columnName))
            ExcelToolsBase.ThrowMissingParameter("columnName", "set-column-number-format");
        if (string.IsNullOrEmpty(formatCode))
            ExcelToolsBase.ThrowMissingParameter("formatCode", "set-column-number-format");

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.SetColumnNumberFormat(batch, tableName!, columnName!, formatCode!);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Column number format set successfully." }, ExcelToolsBase.JsonOptions);
        }
#pragma warning disable CA1031 // MCP protocol requires JSON error responses, not thrown exceptions
        catch (Exception ex)
#pragma warning restore CA1031
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }
}
