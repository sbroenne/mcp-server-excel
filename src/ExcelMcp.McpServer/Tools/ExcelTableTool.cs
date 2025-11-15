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
/// MCP tool for Excel Table (ListObject) operations - structured data with AutoFilter and dynamic expansion.
/// </summary>
[McpServerToolType]
[SuppressMessage("Performance", "CA1861:Avoid constant arrays as arguments", Justification = "Simple workflow arrays in sealed static class")]
public static class TableTool
{
    /// <summary>
    /// Manage Excel Tables (ListObjects) - comprehensive table management including Power Pivot integration
    /// </summary>
    [McpServerTool(Name = "excel_table")]
    [Description(@"Manage Excel Tables (ListObjects) - structured data with AutoFilter")]
    public static async Task<string> Table(
        [Required]
        [Description("Action to perform (enum displayed as dropdown in MCP clients)")]
        TableAction action,

        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")]
        string excelPath,

        [Required]
        [Description("Session ID from excel_file 'open' action")]
        string sessionId,

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
        string? formatCode = null)
    {
        try
        {
            var tableCommands = new TableCommands();

            // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
            return action switch
            {
                TableAction.List => await ListTables(tableCommands, sessionId),
                TableAction.Create => await CreateTable(tableCommands, sessionId, sheetName, tableName, range, hasHeaders, tableStyle),
                TableAction.Get => await GetTableInfo(tableCommands, sessionId, tableName),
                TableAction.Rename => await RenameTable(tableCommands, sessionId, tableName, newName),
                TableAction.Delete => await DeleteTable(tableCommands, sessionId, tableName),
                TableAction.Resize => await ResizeTable(tableCommands, sessionId, tableName, range),
                TableAction.ToggleTotals => await ToggleTotals(tableCommands, sessionId, tableName, hasHeaders),
                TableAction.SetColumnTotal => await SetColumnTotal(tableCommands, sessionId, tableName, newName, tableStyle),
                TableAction.Append => await AppendRows(tableCommands, sessionId, tableName, tableStyle),
                TableAction.SetStyle => await SetTableStyle(tableCommands, sessionId, tableName, tableStyle),
                TableAction.AddToDataModel => await AddToDataModel(tableCommands, sessionId, tableName),
                TableAction.ApplyFilter => await ApplyFilter(tableCommands, sessionId, tableName, newName, filterCriteria),
                TableAction.ApplyFilterValues => await ApplyFilterValues(tableCommands, sessionId, tableName, newName, filterValues),
                TableAction.ClearFilters => await ClearFilters(tableCommands, sessionId, tableName),
                TableAction.GetFilters => await GetFilters(tableCommands, sessionId, tableName),
                TableAction.AddColumn => await AddColumn(tableCommands, sessionId, tableName, newName, filterCriteria),
                TableAction.RemoveColumn => await RemoveColumn(tableCommands, sessionId, tableName, newName),
                TableAction.RenameColumn => await RenameColumn(tableCommands, sessionId, tableName, newName, filterCriteria),
                TableAction.GetStructuredReference => await GetStructuredReference(tableCommands, sessionId, tableName, filterCriteria, newName),
                TableAction.Sort => await SortTable(tableCommands, sessionId, tableName, newName, hasHeaders),
                TableAction.SortMulti => await SortTableMulti(tableCommands, sessionId, tableName, filterValues),
                TableAction.GetColumnNumberFormat => await GetColumnNumberFormat(tableCommands, sessionId, tableName, newName),
                TableAction.SetColumnNumberFormat => await SetColumnNumberFormat(tableCommands, sessionId, tableName, newName, formatCode),
                _ => throw new ArgumentException(
                    $"Unknown action: {action} ({action.ToActionString()})", nameof(action))
            };
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = $"{action.ToActionString()} failed for '{excelPath}': {ex.Message}",
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    private static async Task<string> ListTables(TableCommands commands, string sessionId)
    {
        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ListAsync(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Tables,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateTable(TableCommands commands, string sessionId, string? sheetName, string? tableName, string? range, bool hasHeaders, string? tableStyle)
    {
        if (string.IsNullOrWhiteSpace(sheetName)) ExcelToolsBase.ThrowMissingParameter(nameof(sheetName), "create");
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "create");
        if (string.IsNullOrWhiteSpace(range)) ExcelToolsBase.ThrowMissingParameter(nameof(range), "create");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.CreateAsync(batch, sheetName!, tableName!, range!, hasHeaders, tableStyle));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetTableInfo(TableCommands commands, string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "info");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.GetAsync(batch, tableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Table,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RenameTable(TableCommands commands, string sessionId, string? tableName, string? newName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "rename");
        if (string.IsNullOrWhiteSpace(newName)) ExcelToolsBase.ThrowMissingParameter(nameof(newName), "rename");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.RenameAsync(batch, tableName!, newName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> DeleteTable(TableCommands commands, string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "delete");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.DeleteAsync(batch, tableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ResizeTable(TableCommands commands, string sessionId, string? tableName, string? newRange)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "resize");
        if (string.IsNullOrWhiteSpace(newRange)) ExcelToolsBase.ThrowMissingParameter(nameof(newRange), "resize");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ResizeAsync(batch, tableName!, newRange!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ToggleTotals(TableCommands commands, string sessionId, string? tableName, bool showTotals)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "toggle-totals");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ToggleTotalsAsync(batch, tableName!, showTotals));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetColumnTotal(TableCommands commands, string sessionId, string? tableName, string? columnName, string? totalFunction)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "set-column-total");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "set-column-total");
        if (string.IsNullOrWhiteSpace(totalFunction)) ExcelToolsBase.ThrowMissingParameter(nameof(totalFunction), "set-column-total");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.SetColumnTotalAsync(batch, tableName!, columnName!, totalFunction!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> AppendRows(TableCommands commands, string sessionId, string? tableName, string? csvData)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "append");
        if (string.IsNullOrWhiteSpace(csvData)) ExcelToolsBase.ThrowMissingParameter(nameof(csvData), "append");

        // Parse CSV data to List<List<object?>>
        var rows = ParseCsvToRows(csvData!);

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.AppendAsync(batch, tableName!, rows));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    /// <summary>
    /// Parse CSV data into List of List of objects for table operations.
    /// Simple CSV parser - assumes comma delimiter, handles quoted strings.
    /// </summary>
    private static List<List<object?>> ParseCsvToRows(string csvData)
    {
        var lines = csvData.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);

        var rows = lines.Select(line =>
        {
            var values = line.Split(',');
            return values.Select(value =>
            {
                var trimmed = value.Trim().Trim('"');
                return string.IsNullOrEmpty(trimmed) ? null : (object?)trimmed;
            }).ToList();
        }).ToList();

        return rows;
    }

    private static async Task<string> SetTableStyle(TableCommands commands, string sessionId, string? tableName, string? tableStyle)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "set-style");
        if (string.IsNullOrWhiteSpace(tableStyle)) ExcelToolsBase.ThrowMissingParameter(nameof(tableStyle), "set-style");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.SetStyleAsync(batch, tableName!, tableStyle!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> AddToDataModel(TableCommands commands, string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "add-to-datamodel");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.AddToDataModelAsync(batch, tableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === FILTER OPERATIONS ===

    private static async Task<string> ApplyFilter(TableCommands commands, string sessionId, string? tableName, string? columnName, string? criteria)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "apply-filter");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "apply-filter");
        if (string.IsNullOrWhiteSpace(criteria)) ExcelToolsBase.ThrowMissingParameter(nameof(criteria), "apply-filter");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ApplyFilterAsync(batch, tableName!, columnName!, criteria!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ApplyFilterValues(TableCommands commands, string sessionId, string? tableName, string? columnName, string? filterValuesJson)
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

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ApplyFilterAsync(batch, tableName!, columnName!, filterValues));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ClearFilters(TableCommands commands, string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "clear-filters");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ClearFiltersAsync(batch, tableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetFilters(TableCommands commands, string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "get-filters");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.GetFiltersAsync(batch, tableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.TableName,
            result.ColumnFilters,
            result.HasActiveFilters,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === COLUMN OPERATIONS ===

    private static async Task<string> AddColumn(TableCommands commands, string sessionId, string? tableName, string? columnName, string? positionStr)
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

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.AddColumnAsync(batch, tableName!, columnName!, position));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RemoveColumn(TableCommands commands, string sessionId, string? tableName, string? columnName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "remove-column");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "remove-column");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.RemoveColumnAsync(batch, tableName!, columnName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RenameColumn(TableCommands commands, string sessionId, string? tableName, string? oldColumnName, string? newColumnName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "rename-column");
        if (string.IsNullOrWhiteSpace(oldColumnName)) ExcelToolsBase.ThrowMissingParameter(nameof(oldColumnName), "rename-column");
        if (string.IsNullOrWhiteSpace(newColumnName)) ExcelToolsBase.ThrowMissingParameter(nameof(newColumnName), "rename-column");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.RenameColumnAsync(batch, tableName!, oldColumnName!, newColumnName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === PHASE 2: STRUCTURED REFERENCE & SORT OPERATIONS ===

    private static async Task<string> GetStructuredReference(TableCommands commands, string sessionId, string? tableName, string? regionStr, string? columnName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "get-structured-reference");

        // Parse region string to enum (default: Data)
        var region = Core.Models.TableRegion.Data; // Default
        if (!string.IsNullOrWhiteSpace(regionStr) && !Enum.TryParse(regionStr, true, out region))
        {
            throw new ModelContextProtocol.McpException($"Invalid region '{regionStr}'. Valid values: All, Data, Headers, Totals, ThisRow");
        }

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.GetStructuredReferenceAsync(batch, tableName!, region, columnName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.TableName,
            result.Region,
            result.RangeAddress,
            result.StructuredReference,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SortTable(TableCommands commands, string sessionId, string? tableName, string? columnName, bool ascending)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "sort");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "sort");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.SortAsync(batch, tableName!, columnName!, ascending));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SortTableMulti(TableCommands commands, string sessionId, string? tableName, string? sortColumnsJson)
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

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.SortAsync(batch, tableName!, sortColumns));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === NUMBER FORMAT OPERATIONS ===

    private static async Task<string> GetColumnNumberFormat(TableCommands commands, string sessionId, string? tableName, string? columnName)
    {
        if (string.IsNullOrEmpty(tableName))
            ExcelToolsBase.ThrowMissingParameter("tableName", "get-column-number-format");
        if (string.IsNullOrEmpty(columnName))
            ExcelToolsBase.ThrowMissingParameter("columnName", "get-column-number-format");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.GetColumnNumberFormatAsync(batch, tableName!, columnName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.SheetName,
            result.RangeAddress,
            result.Formats,
            result.RowCount,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetColumnNumberFormat(TableCommands commands, string sessionId, string? tableName, string? columnName, string? formatCode)
    {
        if (string.IsNullOrEmpty(tableName))
            ExcelToolsBase.ThrowMissingParameter("tableName", "set-column-number-format");
        if (string.IsNullOrEmpty(columnName))
            ExcelToolsBase.ThrowMissingParameter("columnName", "set-column-number-format");
        if (string.IsNullOrEmpty(formatCode))
            ExcelToolsBase.ThrowMissingParameter("formatCode", "set-column-number-format");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.SetColumnNumberFormatAsync(batch, tableName!, columnName!, formatCode!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }
}
