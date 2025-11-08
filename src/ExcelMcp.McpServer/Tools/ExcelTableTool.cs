using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics.CodeAnalysis;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.McpServer.Models;

#pragma warning disable CA1861 // Avoid constant arrays as arguments - workflow hints are contextual per-call

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel Table (ListObject) management tool for MCP server.
/// Handles creating, listing, renaming, and deleting Excel Tables.
///
/// LLM Usage Patterns:
/// - Use "list" to see all Excel Tables in a workbook
/// - Use "create" to convert ranges to Excel Tables (enables AutoFilter, structured references, dynamic expansion)
/// - Use "info" to get detailed information about a table
/// - Use "rename" to change table names
/// - Use "delete" to remove tables (converts back to range, data preserved)
/// - Use "add-to-datamodel" to add a table to Power Pivot Data Model
///
/// IMPORTANT:
/// - Excel Tables provide AutoFilter, structured references ([@Column]), dynamic expansion, and visual formatting
/// - Tables can be used standalone OR referenced in Power Query: Excel.CurrentWorkbook(){[Name="TableName"]}[Content]
/// - Table names must start with a letter/underscore, contain only alphanumeric and underscore characters
/// - Column names can be any string, including purely numeric values (e.g., "60" for 60 months)
/// - Deleting a table converts it back to a range but preserves data
/// - For comprehensive Power Pivot operations (DAX measures, relationships, calculated columns), use excel_datamodel tools
/// </summary>
[McpServerToolType]
[SuppressMessage("Performance", "CA1861:Avoid constant arrays as arguments", Justification = "Simple workflow arrays in sealed static class")]
public static class TableTool
{
    /// <summary>
    /// Manage Excel Tables (ListObjects) - comprehensive table management including Power Pivot integration
    /// </summary>
    [McpServerTool(Name = "excel_table")]
    [Description("Manage Excel Tables (ListObjects). Tables provide AutoFilter, structured references, dynamic expansion, and visual formatting. Can be used standalone or referenced in Power Query. Use 'add-to-datamodel' action to add existing tables to Power Pivot. For Power Pivot workflows: create table here → use excel_powerquery to load external data to Power Pivot → use excel_datamodel/excel_powerpivot for DAX measures and relationships. Supports: list, create, info, rename, delete, resize, toggle-totals, set-column-total, append, set-style, add-to-datamodel, apply-filter, apply-filter-values, clear-filters, get-filters, add-column, remove-column, rename-column, get-structured-reference, sort, sort-multi, get-column-number-format, set-column-number-format.")]
    public static async Task<string> Table(
        [Required]
        [Description("Action to perform (enum displayed as dropdown in MCP clients)")]
        TableAction action,

        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")]
        string excelPath,

        [StringLength(255, MinimumLength = 1)]
        [RegularExpression(@"^[a-zA-Z_][a-zA-Z0-9_]*$")]
        [Description("Table name (required for most actions). Must start with letter/underscore, alphanumeric + underscore only")]
        string? tableName = null,

        [StringLength(31, MinimumLength = 1)]
        [RegularExpression(@"^[^[\]/*?\\:]+$")]
        [Description("Sheet name (required for create)")]
        string? sheetName = null,

        [Description("Excel range (e.g., 'A1:D10') - required for create/resize")]
        string? range = null,

        [StringLength(255, MinimumLength = 1)]
        [Description("New table name (required for rename) or column name (required for add-column, rename-column, etc.). Table names must follow Excel naming rules (start with letter/underscore, alphanumeric only). Column names can be any string including numbers.")]
        string? newName = null,

        [Description("Whether the range has headers (default: true for create) or show totals (for toggle-totals)")]
        bool hasHeaders = true,

        [Description("Table style name (e.g., 'TableStyleMedium2') for create/set-style, or total function (sum/avg/count) for set-column-total, or CSV data for append")]
        string? tableStyle = null,

        [Description("Filter criteria (e.g., '>100', '=Text') for apply-filter, or column position (0-based) for add-column")]
        string? filterCriteria = null,

        [Description("JSON array of filter values (e.g., '[\"Value1\",\"Value2\"]') for apply-filter-values")]
        string? filterValues = null,

        [Description("Excel format code for set-column-number-format (e.g., '$#,##0.00', '0.00%', 'm/d/yyyy')")]
        string? formatCode = null,

        [Description("Optional batch session ID from begin_excel_batch (for multi-operation workflows)")]
        string? batchId = null)
    {
        try
        {
            var tableCommands = new TableCommands();

            // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
            return action switch
            {
                TableAction.List => await ListTables(tableCommands, excelPath, batchId),
                TableAction.Create => await CreateTable(tableCommands, excelPath, sheetName, tableName, range, hasHeaders, tableStyle, batchId),
                TableAction.Get => await GetTableInfo(tableCommands, excelPath, tableName, batchId),
                TableAction.Rename => await RenameTable(tableCommands, excelPath, tableName, newName, batchId),
                TableAction.Delete => await DeleteTable(tableCommands, excelPath, tableName, batchId),
                TableAction.Resize => await ResizeTable(tableCommands, excelPath, tableName, range, batchId),
                TableAction.ToggleTotals => await ToggleTotals(tableCommands, excelPath, tableName, hasHeaders, batchId),
                TableAction.SetColumnTotal => await SetColumnTotal(tableCommands, excelPath, tableName, newName, tableStyle, batchId),
                TableAction.Append => await AppendRows(tableCommands, excelPath, tableName, tableStyle, batchId),
                TableAction.SetStyle => await SetTableStyle(tableCommands, excelPath, tableName, tableStyle, batchId),
                TableAction.AddToDataModel => await AddToDataModel(tableCommands, excelPath, tableName, batchId),
                TableAction.ApplyFilter => await ApplyFilter(tableCommands, excelPath, tableName, newName, filterCriteria, batchId),
                TableAction.ApplyFilterValues => await ApplyFilterValues(tableCommands, excelPath, tableName, newName, filterValues, batchId),
                TableAction.ClearFilters => await ClearFilters(tableCommands, excelPath, tableName, batchId),
                TableAction.GetFilters => await GetFilters(tableCommands, excelPath, tableName, batchId),
                TableAction.AddColumn => await AddColumn(tableCommands, excelPath, tableName, newName, filterCriteria, batchId),
                TableAction.RemoveColumn => await RemoveColumn(tableCommands, excelPath, tableName, newName, batchId),
                TableAction.RenameColumn => await RenameColumn(tableCommands, excelPath, tableName, newName, filterCriteria, batchId),
                TableAction.GetStructuredReference => await GetStructuredReference(tableCommands, excelPath, tableName, filterCriteria, newName, batchId),
                TableAction.Sort => await SortTable(tableCommands, excelPath, tableName, newName, hasHeaders, batchId),
                TableAction.SortMulti => await SortTableMulti(tableCommands, excelPath, tableName, filterValues, batchId),
                TableAction.GetColumnNumberFormat => await GetColumnNumberFormat(tableCommands, excelPath, tableName, newName, batchId),
                TableAction.SetColumnNumberFormat => await SetColumnNumberFormat(tableCommands, excelPath, tableName, newName, formatCode, batchId),
                _ => throw new ModelContextProtocol.McpException(
                    $"Unknown action: {action} ({action.ToActionString()})")
            };
        }
        catch (ModelContextProtocol.McpException)
        {
            throw; // Re-throw MCP exceptions as-is
        }
        catch (Exception ex)
        {
            ExcelToolsBase.ThrowInternalError(ex, action.ToActionString(), excelPath);
            throw; // Unreachable but satisfies compiler
        }
    }

    private static async Task<string> ListTables(TableCommands commands, string filePath, string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            false, // don't save for list operation
            async (batch) => await commands.ListAsync(batch)
        );

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Tables,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Found {result.Tables.Count} Excel Tables. Use for structured data with AutoFilter and dynamic expansion."
                : "Failed to list tables. Verify workbook contains Excel Tables (ListObjects).",
            suggestedNextActions = result.Success
                ? new[] { "Use 'get' to view detailed table information", "Use 'create' to convert ranges to Excel Tables", "Use excel_range to read table data" }
                : ["Create tables with 'create' action", "Verify workbook has data ranges", "Check if workbook is corrupted"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateTable(TableCommands commands, string filePath, string? sheetName, string? tableName, string? range, bool hasHeaders, string? tableStyle, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(sheetName)) ExcelToolsBase.ThrowMissingParameter(nameof(sheetName), "create");
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "create");
        if (string.IsNullOrWhiteSpace(range)) ExcelToolsBase.ThrowMissingParameter(nameof(range), "create");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            true, // save changes
            async (batch) => await commands.CreateAsync(batch, sheetName!, tableName!, range!, hasHeaders, tableStyle)
        );

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Table '{tableName}' created successfully. Data now has AutoFilter and structured references enabled."
                : $"Failed to create table '{tableName}'. Verify range exists and doesn't overlap existing tables.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'get' to view table details", "Use excel_range to populate table data", "Use 'set-style' to apply table formatting" }
                : ["Verify range address is valid (e.g., A1:D10)", "Check sheet name is correct", "Ensure range doesn't overlap existing tables"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetTableInfo(TableCommands commands, string filePath, string? tableName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "info");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            false, // don't save for info operation
            async (batch) => await commands.GetAsync(batch, tableName!)
        );

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Table,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Table '{tableName}' details retrieved. Review columns, range, and configuration."
                : $"Failed to get info for table '{tableName}'. Verify table name is correct.",
            suggestedNextActions = result.Success
                ? new[] { "Use excel_range to read table data", "Use 'resize' to adjust table range", "Use 'toggle-totals' to enable summary row" }
                : ["Use 'list' to see all available table names", "Check for typos in table name", "Verify table exists in workbook"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RenameTable(TableCommands commands, string filePath, string? tableName, string? newName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "rename");
        if (string.IsNullOrWhiteSpace(newName)) ExcelToolsBase.ThrowMissingParameter(nameof(newName), "rename");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            true, // save changes
            async (batch) => await commands.RenameAsync(batch, tableName!, newName!)
        );

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Table renamed from '{tableName}' to '{newName}'. All structured references updated automatically."
                : $"Failed to rename table '{tableName}'. Verify new name doesn't conflict with existing tables.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'list' to verify rename", "Use 'get' to view updated table info", "Update formulas using old name if needed" }
                : ["Check new name doesn't already exist", "Verify table name format (no spaces, special chars)", "Ensure table exists in workbook"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> DeleteTable(TableCommands commands, string filePath, string? tableName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "delete");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            true, // save changes
            async (batch) => await commands.DeleteAsync(batch, tableName!)
        );

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Table '{tableName}' deleted. Converted back to range - data preserved but AutoFilter and structured references removed."
                : $"Failed to delete table '{tableName}'. Verify table exists.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'list' to verify deletion", "Use excel_range to access remaining data as range", "Create new table if needed with 'create'" }
                : ["Use 'list' to verify table name", "Check if table is in use by formulas", "Ensure table exists in workbook"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ResizeTable(TableCommands commands, string filePath, string? tableName, string? newRange, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "resize");
        if (string.IsNullOrWhiteSpace(newRange)) ExcelToolsBase.ThrowMissingParameter(nameof(newRange), "resize");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            true, // save changes
            async (batch) => await commands.ResizeAsync(batch, tableName!, newRange!)
        );

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Table '{tableName}' resized to {newRange}. Data and formulas adjusted automatically."
                : $"Failed to resize table '{tableName}'. Verify new range is valid and doesn't overlap other tables.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'get' to verify new table dimensions", "Use excel_range to populate new rows/columns", "Use 'append' to add data to expanded table" }
                : ["Check new range format (e.g., A1:D20)", "Verify range doesn't overlap existing tables", "Ensure range includes all existing data"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ToggleTotals(TableCommands commands, string filePath, string? tableName, bool showTotals, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "toggle-totals");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            true, // save changes
            async (batch) => await commands.ToggleTotalsAsync(batch, tableName!, showTotals)
        );

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Totals row {(showTotals ? "enabled" : "disabled")} for table '{tableName}'."
                : $"Failed to toggle totals for table '{tableName}'. Verify table exists.",
            suggestedNextActions = result.Success
                ? (showTotals
                    ? new[] { "Use 'set-column-total' to configure aggregation functions", "Use excel_range to read totals row values", "Use 'get' to view updated table info" }
                    : ["Use 'toggle-totals' with true to re-enable totals", "Use 'get' to verify totals row removed", "Use excel_range to read table data"])
                : ["Use 'list' to verify table name", "Check if table exists", "Verify table has data rows"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetColumnTotal(TableCommands commands, string filePath, string? tableName, string? columnName, string? totalFunction, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "set-column-total");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "set-column-total");
        if (string.IsNullOrWhiteSpace(totalFunction)) ExcelToolsBase.ThrowMissingParameter(nameof(totalFunction), "set-column-total");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            true, // save changes
            async (batch) => await commands.SetColumnTotalAsync(batch, tableName!, columnName!, totalFunction!)
        );

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Total function '{totalFunction}' applied to column '{columnName}' in table '{tableName}'."
                : $"Failed to set column total for '{columnName}'. Verify column exists and totals row is enabled.",
            suggestedNextActions = result.Success
                ? new[] { "Use excel_range to read totals row values", "Use 'get' to view updated table info", "Use 'set-column-total' for other columns" }
                : ["Use 'toggle-totals' to enable totals row first", "Verify column name is correct", "Check totalFunction is valid (sum, avg, count, max, min)"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> AppendRows(TableCommands commands, string filePath, string? tableName, string? csvData, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "append");
        if (string.IsNullOrWhiteSpace(csvData)) ExcelToolsBase.ThrowMissingParameter(nameof(csvData), "append");

        // Parse CSV data to List<List<object?>>
        var rows = ParseCsvToRows(csvData!);

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            true, // save changes
            async (batch) => await commands.AppendAsync(batch, tableName!, rows)
        );

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Rows appended to table '{tableName}'. Table automatically expanded to include new data."
                : $"Failed to append rows to table '{tableName}'. Verify CSV format matches table columns.",
            suggestedNextActions = result.Success
                ? new[] { "Use excel_range to read updated table data", "Use 'get' to verify new table dimensions", "Use 'resize' to adjust if needed" }
                : ["Verify CSV column count matches table columns", "Check CSV format (comma-separated values)", "Ensure table exists in workbook"]
        }, ExcelToolsBase.JsonOptions);
    }

    /// <summary>
    /// Parse CSV data into List of List of objects for table operations.
    /// Simple CSV parser - assumes comma delimiter, handles quoted strings.
    /// </summary>
    private static List<List<object?>> ParseCsvToRows(string csvData)
    {
        var rows = new List<List<object?>>();
        var lines = csvData.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);

        foreach (var line in lines)
        {
            var values = line.Split(',');
            var row = values.Select(value =>
            {
                var trimmed = value.Trim().Trim('"');
                return string.IsNullOrEmpty(trimmed) ? null : (object?)trimmed;
            }).ToList();

            rows.Add(row);
        }

        return rows;
    }

    private static async Task<string> SetTableStyle(TableCommands commands, string filePath, string? tableName, string? tableStyle, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "set-style");
        if (string.IsNullOrWhiteSpace(tableStyle)) ExcelToolsBase.ThrowMissingParameter(nameof(tableStyle), "set-style");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            true, // save changes
            async (batch) => await commands.SetStyleAsync(batch, tableName!, tableStyle!)
        );

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Table style '{tableStyle}' applied to '{tableName}'. Visual formatting updated."
                : $"Failed to set style for table '{tableName}'. Verify style name is valid.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'get' to view updated table appearance", "Use excel_range to read formatted data", "Use 'create' with different style for new tables" }
                : ["Check style name (e.g., TableStyleMedium2, TableStyleLight1)", "Use 'list' to verify table exists", "Try different predefined table style"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> AddToDataModel(TableCommands commands, string filePath, string? tableName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "add-to-datamodel");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            true, // save changes
            async (batch) => await commands.AddToDataModelAsync(batch, tableName!)
        );

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Table '{tableName}' added to Power Pivot Data Model. Ready for DAX measures and relationships."
                : $"Failed to add table '{tableName}' to Data Model. Verify table exists and workbook supports Data Model.",
            suggestedNextActions = result.Success
                ? new[] { "Use excel_datamodel 'list-tables' to verify table in model", "Use excel_datamodel 'create-measure' to add DAX calculations", "Use excel_datamodel 'create-relationship' to connect tables" }
                : ["Verify workbook is .xlsx format (not .xls)", "Check if Power Pivot is available", "Use 'list' to verify table exists"]
        }, ExcelToolsBase.JsonOptions);
    }

    // === FILTER OPERATIONS ===

    private static async Task<string> ApplyFilter(TableCommands commands, string filePath, string? tableName, string? columnName, string? criteria, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "apply-filter");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "apply-filter");
        if (string.IsNullOrWhiteSpace(criteria)) ExcelToolsBase.ThrowMissingParameter(nameof(criteria), "apply-filter");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            true,
            async (batch) => await commands.ApplyFilterAsync(batch, tableName!, columnName!, criteria!)
        );

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Filter applied to column '{columnName}' in table '{tableName}' with criteria '{criteria}'."
                : $"Failed to apply filter to column '{columnName}'. Verify column exists and criteria format is valid.",
            suggestedNextActions = result.Success
                ? new[] { "Use excel_range to read filtered data", "Use 'get-filters' to view active filters", "Use 'clear-filters' to remove all filters" }
                : ["Verify column name is correct", "Check criteria format (e.g., >100, =Active, <>Closed)", "Use 'get' to see available columns"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ApplyFilterValues(TableCommands commands, string filePath, string? tableName, string? columnName, string? filterValuesJson, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "apply-filter-values");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "apply-filter-values");
        if (string.IsNullOrWhiteSpace(filterValuesJson)) ExcelToolsBase.ThrowMissingParameter(nameof(filterValuesJson), "apply-filter-values");

        // Parse JSON array to List<string>
        List<string> filterValues;
        try
        {
            filterValues = JsonSerializer.Deserialize<List<string>>(filterValuesJson!) ?? [];
        }
        catch (JsonException ex)
        {
            throw new ModelContextProtocol.McpException($"Invalid JSON array for filterValues: {ex.Message}");
        }

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            true,
            async (batch) => await commands.ApplyFilterAsync(batch, tableName!, columnName!, filterValues)
        );

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Filter applied to column '{columnName}' with {filterValues.Count} specific values."
                : $"Failed to apply value filter to column '{columnName}'. Verify column exists and values are valid.",
            suggestedNextActions = result.Success
                ? new[] { "Use excel_range to read filtered data", "Use 'get-filters' to view active filters", "Use 'apply-filter' for criteria-based filtering" }
                : ["Verify column name is correct", "Check filterValues format (JSON array of strings)", "Use 'get' to see available columns"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ClearFilters(TableCommands commands, string filePath, string? tableName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "clear-filters");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            true,
            async (batch) => await commands.ClearFiltersAsync(batch, tableName!)
        );

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"All filters cleared from table '{tableName}'. Full dataset is now visible."
                : $"Failed to clear filters from table '{tableName}'.",
            suggestedNextActions = result.Success
                ? new[] { "Use excel_range to read all table data", "Use 'apply-filter' to set new filters", "Use 'get' to view table information" }
                : ["Verify table exists with 'list'", "Check if table has filters with 'get-filters'", "Ensure workbook is not protected"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetFilters(TableCommands commands, string filePath, string? tableName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "get-filters");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            false,
            async (batch) => await commands.GetFiltersAsync(batch, tableName!)
        );

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.TableName,
            result.ColumnFilters,
            result.HasActiveFilters,
            result.ErrorMessage,
            workflowHint = result.Success
                ? (result.HasActiveFilters
                    ? $"Table '{tableName}' has {result.ColumnFilters?.Count ?? 0} active filter(s)."
                    : $"Table '{tableName}' has no active filters.")
                : $"Failed to retrieve filters from table '{tableName}'.",
            suggestedNextActions = result.Success
                ? (result.HasActiveFilters
                    ? new[] { "Use 'clear-filters' to remove all filters", "Use excel_range to read filtered data", "Use 'apply-filter' to modify filter criteria" }
                    : ["Use 'apply-filter' to filter by criteria", "Use 'apply-filter-values' to filter by specific values", "Use excel_range to read all table data"])
                : ["Verify table exists with 'list'", "Check table name spelling", "Ensure table has AutoFilter enabled"]
        }, ExcelToolsBase.JsonOptions);
    }

    // === COLUMN OPERATIONS ===

    private static async Task<string> AddColumn(TableCommands commands, string filePath, string? tableName, string? columnName, string? positionStr, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "add-column");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "add-column");

        // Parse position (optional)
        int? position = null;
        if (!string.IsNullOrWhiteSpace(positionStr))
        {
            if (int.TryParse(positionStr, out int pos))
            {
                position = pos;
            }
            else
            {
                throw new ModelContextProtocol.McpException($"Invalid position value: '{positionStr}'. Must be a number.");
            }
        }

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            true,
            async (batch) => await commands.AddColumnAsync(batch, tableName!, columnName!, position)
        );

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Column '{columnName}' added to table '{tableName}'{(position.HasValue ? $" at position {position.Value}" : "")}."
                : $"Failed to add column '{columnName}' to table '{tableName}'.",
            suggestedNextActions = result.Success
                ? new[] { "Use excel_range to populate the new column with data", "Use 'rename-column' to change column name if needed", "Use 'get' to view updated table structure" }
                : ["Verify table exists with 'list'", "Check if column name already exists", "Ensure position is within table bounds"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RemoveColumn(TableCommands commands, string filePath, string? tableName, string? columnName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "remove-column");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "remove-column");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            true,
            async (batch) => await commands.RemoveColumnAsync(batch, tableName!, columnName!)
        );

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Column '{columnName}' removed from table '{tableName}'. Table structure updated."
                : $"Failed to remove column '{columnName}' from table '{tableName}'.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'get' to view updated table structure", "Use 'list' to verify table still exists", "Use excel_range to verify remaining data" }
                : ["Verify column name is correct with 'get'", "Check if column is the last column in table", "Ensure table has more than one column"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RenameColumn(TableCommands commands, string filePath, string? tableName, string? oldColumnName, string? newColumnName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "rename-column");
        if (string.IsNullOrWhiteSpace(oldColumnName)) ExcelToolsBase.ThrowMissingParameter(nameof(oldColumnName), "rename-column");
        if (string.IsNullOrWhiteSpace(newColumnName)) ExcelToolsBase.ThrowMissingParameter(nameof(newColumnName), "rename-column");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            true,
            async (batch) => await commands.RenameColumnAsync(batch, tableName!, oldColumnName!, newColumnName!)
        );

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Column '{oldColumnName}' renamed to '{newColumnName}' in table '{tableName}'. Formulas referencing this column will update automatically."
                : $"Failed to rename column '{oldColumnName}' to '{newColumnName}' in table '{tableName}'.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'get' to verify column name change", "Use 'get-structured-reference' to see updated column reference", "Update any external references to this column" }
                : ["Verify old column name is correct with 'get'", "Check if new column name already exists", "Ensure new column name follows Excel naming rules"]
        }, ExcelToolsBase.JsonOptions);
    }

    // === PHASE 2: STRUCTURED REFERENCE & SORT OPERATIONS ===

    private static async Task<string> GetStructuredReference(TableCommands commands, string filePath, string? tableName, string? regionStr, string? columnName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "get-structured-reference");

        // Parse region string to enum (default: Data)
        var region = Core.Models.TableRegion.Data; // Default
        if (!string.IsNullOrWhiteSpace(regionStr))
        {
            if (!Enum.TryParse(regionStr, true, out region))
            {
                throw new ModelContextProtocol.McpException($"Invalid region '{regionStr}'. Valid values: All, Data, Headers, Totals, ThisRow");
            }
        }

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            false, // Read-only operation
            async (batch) => await commands.GetStructuredReferenceAsync(batch, tableName!, region, columnName)
        );

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.TableName,
            result.Region,
            result.RangeAddress,
            result.StructuredReference,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Structured reference for table '{tableName}', region '{region}'{(!string.IsNullOrEmpty(columnName) ? $", column '{columnName}'" : "")} is '{result.StructuredReference}'."
                : $"Failed to get structured reference for table '{tableName}'.",
            suggestedNextActions = result.Success
                ? new[] { "Use structured reference in Excel formulas", "Use excel_range with RangeAddress to read/write data", "Use different region to get other table parts" }
                : ["Verify table exists with 'list'", "Check region value (All, Data, Headers, Totals, ThisRow)", "Verify column name if specified"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SortTable(TableCommands commands, string filePath, string? tableName, string? columnName, bool ascending, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "sort");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "sort");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            true, // Save changes
            async (batch) => await commands.SortAsync(batch, tableName!, columnName!, ascending)
        );

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Table '{tableName}' sorted by column '{columnName}' in {(ascending ? "ascending" : "descending")} order."
                : $"Failed to sort table '{tableName}' by column '{columnName}'.",
            suggestedNextActions = result.Success
                ? new[] { "Use excel_range to read sorted data", "Use 'sort-multi' for multi-level sorting", "Use 'get' to view table information" }
                : ["Verify column name is correct with 'get'", "Check if table has data to sort", "Ensure table is not protected"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SortTableMulti(TableCommands commands, string filePath, string? tableName, string? sortColumnsJson, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "sort-multi");
        if (string.IsNullOrWhiteSpace(sortColumnsJson)) ExcelToolsBase.ThrowMissingParameter(nameof(sortColumnsJson), "sort-multi");

        // Parse JSON array of sort columns
        List<Core.Models.TableSortColumn>? sortColumns;
        try
        {
            sortColumns = JsonSerializer.Deserialize<List<Core.Models.TableSortColumn>>(sortColumnsJson!);
            if (sortColumns == null || sortColumns.Count == 0)
            {
                throw new ModelContextProtocol.McpException("sortColumns JSON must be a non-empty array");
            }
        }
        catch (JsonException ex)
        {
            throw new ModelContextProtocol.McpException($"Invalid sortColumns JSON: {ex.Message}");
        }

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            true, // Save changes
            async (batch) => await commands.SortAsync(batch, tableName!, sortColumns)
        );

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Table '{tableName}' sorted by {sortColumns.Count} column(s) with multi-level criteria."
                : $"Failed to sort table '{tableName}' with multiple columns.",
            suggestedNextActions = result.Success
                ? new[] { "Use excel_range to read sorted data", "Use 'get' to view table information", "Use 'sort' for single-column sorting" }
                : ["Verify all column names exist with 'get'", "Check sortColumns JSON format", "Ensure table has data to sort"]
        }, ExcelToolsBase.JsonOptions);
    }

    // === NUMBER FORMAT OPERATIONS ===

    private static async Task<string> GetColumnNumberFormat(TableCommands commands, string filePath, string? tableName, string? columnName, string? batchId)
    {
        if (string.IsNullOrEmpty(tableName))
            ExcelToolsBase.ThrowMissingParameter("tableName", "get-column-number-format");
        if (string.IsNullOrEmpty(columnName))
            ExcelToolsBase.ThrowMissingParameter("columnName", "get-column-number-format");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.GetColumnNumberFormatAsync(batch, tableName!, columnName!)
        );

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.SheetName,
            result.RangeAddress,
            result.Formats,
            result.RowCount,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Retrieved number format for column '{columnName}' in table '{tableName}' ({result.RowCount} rows)."
                : $"Failed to get number format for column '{columnName}' in table '{tableName}'.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'set-column-number-format' to change format", "Use excel_range to read formatted values", "Use 'get' to view table structure" }
                : ["Verify column name is correct with 'get'", "Check if table exists with 'list'", "Ensure table has data"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetColumnNumberFormat(TableCommands commands, string filePath, string? tableName, string? columnName, string? formatCode, string? batchId)
    {
        if (string.IsNullOrEmpty(tableName))
            ExcelToolsBase.ThrowMissingParameter("tableName", "set-column-number-format");
        if (string.IsNullOrEmpty(columnName))
            ExcelToolsBase.ThrowMissingParameter("columnName", "set-column-number-format");
        if (string.IsNullOrEmpty(formatCode))
            ExcelToolsBase.ThrowMissingParameter("formatCode", "set-column-number-format");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.SetColumnNumberFormatAsync(batch, tableName!, columnName!, formatCode!)
        );

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Number format '{formatCode}' applied to column '{columnName}' in table '{tableName}'."
                : $"Failed to set number format for column '{columnName}' in table '{tableName}'.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'get-column-number-format' to verify format applied", "Use excel_range to read formatted values", "Use 'get' to view table structure" }
                : ["Verify column name is correct with 'get'", "Check format code syntax (e.g., '#,##0.00', '0.00%')", "Ensure table exists with 'list'"]
        }, ExcelToolsBase.JsonOptions);
    }
}
