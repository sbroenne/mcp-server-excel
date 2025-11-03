using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.McpServer.Models;

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
/// - Deleting a table converts it back to a range but preserves data
/// - For comprehensive Power Pivot operations (DAX measures, relationships, calculated columns), use excel_datamodel tools
/// </summary>
[McpServerToolType]
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
        [RegularExpression(@"^[a-zA-Z_][a-zA-Z0-9_]*$")]
        [Description("New table name (required for rename) or column name (required for set-column-total)")]
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {

            throw new ModelContextProtocol.McpException($"list failed for '{filePath}': {result.ErrorMessage}");
        }

        if (result.Tables == null || !result.Tables.Any())
        {

        }
        else
        {

        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {

            throw new ModelContextProtocol.McpException($"create failed for table '{tableName}': {result.ErrorMessage}");
        }





        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {

            throw new ModelContextProtocol.McpException($"info failed for table '{tableName}': {result.ErrorMessage}");
        }





        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {

            throw new ModelContextProtocol.McpException($"rename failed for table '{tableName}': {result.ErrorMessage}");
        }





        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {

            throw new ModelContextProtocol.McpException($"delete failed for table '{tableName}': {result.ErrorMessage}");
        }





        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"resize failed for table '{tableName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"toggle-totals failed for table '{tableName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set-column-total failed for table '{tableName}', column '{columnName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"append failed for table '{tableName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    /// <summary>
    /// Parse CSV data into List of List of objects for table operations.
    /// Simple CSV parser - assumes comma delimiter, handles quoted strings.
    /// </summary>
    private static List<List<object?>> ParseCsvToRows(string csvData)
    {
        var rows = new List<List<object?>>();
        var lines = csvData.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

        foreach (var line in lines)
        {
            var row = new List<object?>();
            var values = line.Split(',');

            foreach (var value in values)
            {
                var trimmed = value.Trim().Trim('"');
                row.Add(string.IsNullOrEmpty(trimmed) ? null : trimmed);
            }

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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set-style failed for table '{tableName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"add-to-datamodel failed for table '{tableName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"apply-filter failed for table '{tableName}', column '{columnName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"apply-filter-values failed for table '{tableName}', column '{columnName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"clear-filters failed for table '{tableName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"get-filters failed for table '{tableName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"add-column failed for table '{tableName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"remove-column failed for table '{tableName}', column '{columnName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"rename-column failed for table '{tableName}', column '{oldColumnName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    // === PHASE 2: STRUCTURED REFERENCE & SORT OPERATIONS ===

    private static async Task<string> GetStructuredReference(TableCommands commands, string filePath, string? tableName, string? regionStr, string? columnName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "get-structured-reference");

        // Parse region string to enum (default: Data)
        var region = Core.Models.TableRegion.Data; // Default
        if (!string.IsNullOrWhiteSpace(regionStr))
        {
            if (!Enum.TryParse<Core.Models.TableRegion>(regionStr, true, out region))
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"get-structured-reference failed for table '{tableName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"sort failed for table '{tableName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"sort-multi failed for table '{tableName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"get-column-number-format failed for table '{tableName}' column '{columnName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set-column-number-format failed for table '{tableName}' column '{columnName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}
