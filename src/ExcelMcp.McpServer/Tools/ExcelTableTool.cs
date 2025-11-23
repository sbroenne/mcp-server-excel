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
    public static string Table(
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
        return ExcelToolsBase.ExecuteToolAction(
            action.ToActionString(),
            excelPath,
            () =>
            {
                var tableCommands = new TableCommands();

                // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
                return action switch
                {
                    TableAction.List => ListTables(tableCommands, sessionId),
                    TableAction.Create => CreateTable(tableCommands, sessionId, sheetName, tableName, range, hasHeaders, tableStyle),
                    TableAction.Read => ReadTable(tableCommands, sessionId, tableName),
                    TableAction.Rename => RenameTable(tableCommands, sessionId, tableName, newName),
                    TableAction.Delete => DeleteTable(tableCommands, sessionId, tableName),
                    TableAction.Resize => ResizeTable(tableCommands, sessionId, tableName, range),
                    TableAction.ToggleTotals => ToggleTotals(tableCommands, sessionId, tableName, hasHeaders),
                    TableAction.SetColumnTotal => SetColumnTotal(tableCommands, sessionId, tableName, newName, tableStyle),
                    TableAction.Append => AppendRows(tableCommands, sessionId, tableName, tableStyle),
                    TableAction.SetStyle => SetTableStyle(tableCommands, sessionId, tableName, tableStyle),
                    TableAction.AddToDataModel => AddToDataModel(tableCommands, sessionId, tableName),
                    TableAction.ApplyFilter => ApplyFilter(tableCommands, sessionId, tableName, newName, filterCriteria),
                    TableAction.ApplyFilterValues => ApplyFilterValues(tableCommands, sessionId, tableName, newName, filterValues),
                    TableAction.ClearFilters => ClearFilters(tableCommands, sessionId, tableName),
                    TableAction.GetFilters => GetFilters(tableCommands, sessionId, tableName),
                    TableAction.AddColumn => AddColumn(tableCommands, sessionId, tableName, newName, filterCriteria),
                    TableAction.RemoveColumn => RemoveColumn(tableCommands, sessionId, tableName, newName),
                    TableAction.RenameColumn => RenameColumn(tableCommands, sessionId, tableName, newName, filterCriteria),
                    TableAction.GetStructuredReference => GetStructuredReference(tableCommands, sessionId, tableName, filterCriteria, newName),
                    TableAction.Sort => SortTable(tableCommands, sessionId, tableName, newName, hasHeaders),
                    TableAction.SortMulti => SortTableMulti(tableCommands, sessionId, tableName, filterValues),
                    TableAction.GetColumnNumberFormat => GetColumnNumberFormat(tableCommands, sessionId, tableName, newName),
                    TableAction.SetColumnNumberFormat => SetColumnNumberFormat(tableCommands, sessionId, tableName, newName, formatCode),
                    _ => throw new ArgumentException(
                        $"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string ListTables(TableCommands commands, string sessionId)
    {
        var result = ExcelToolsBase.WithSession(sessionId, batch => commands.List(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Tables,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string CreateTable(TableCommands commands, string sessionId, string? sheetName, string? tableName, string? range, bool hasHeaders, string? tableStyle)
    {
        if (string.IsNullOrWhiteSpace(sheetName)) ExcelToolsBase.ThrowMissingParameter(nameof(sheetName), "create");
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "create");
        if (string.IsNullOrWhiteSpace(range)) ExcelToolsBase.ThrowMissingParameter(nameof(range), "create");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Create(batch, sheetName!, tableName!, range!, hasHeaders, tableStyle));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ReadTable(TableCommands commands, string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "read");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Read(batch, tableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Table,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string RenameTable(TableCommands commands, string sessionId, string? tableName, string? newName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "rename");
        if (string.IsNullOrWhiteSpace(newName)) ExcelToolsBase.ThrowMissingParameter(nameof(newName), "rename");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Rename(batch, tableName!, newName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteTable(TableCommands commands, string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "delete");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Delete(batch, tableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ResizeTable(TableCommands commands, string sessionId, string? tableName, string? newRange)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "resize");
        if (string.IsNullOrWhiteSpace(newRange)) ExcelToolsBase.ThrowMissingParameter(nameof(newRange), "resize");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Resize(batch, tableName!, newRange!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ToggleTotals(TableCommands commands, string sessionId, string? tableName, bool showTotals)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "toggle-totals");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.ToggleTotals(batch, tableName!, showTotals));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string SetColumnTotal(TableCommands commands, string sessionId, string? tableName, string? columnName, string? totalFunction)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "set-column-total");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "set-column-total");
        if (string.IsNullOrWhiteSpace(totalFunction)) ExcelToolsBase.ThrowMissingParameter(nameof(totalFunction), "set-column-total");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.SetColumnTotal(batch, tableName!, columnName!, totalFunction!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string AppendRows(TableCommands commands, string sessionId, string? tableName, string? csvData)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "append");
        if (string.IsNullOrWhiteSpace(csvData)) ExcelToolsBase.ThrowMissingParameter(nameof(csvData), "append");

        // Parse CSV data to List<List<object?>>
        var rows = ParseCsvToRows(csvData!);

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Append(batch, tableName!, rows));

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

    private static string SetTableStyle(TableCommands commands, string sessionId, string? tableName, string? tableStyle)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "set-style");
        if (string.IsNullOrWhiteSpace(tableStyle)) ExcelToolsBase.ThrowMissingParameter(nameof(tableStyle), "set-style");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.SetStyle(batch, tableName!, tableStyle!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string AddToDataModel(TableCommands commands, string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "add-to-datamodel");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.AddToDataModel(batch, tableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === FILTER OPERATIONS ===

    private static string ApplyFilter(TableCommands commands, string sessionId, string? tableName, string? columnName, string? criteria)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "apply-filter");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "apply-filter");
        if (string.IsNullOrWhiteSpace(criteria)) ExcelToolsBase.ThrowMissingParameter(nameof(criteria), "apply-filter");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.ApplyFilter(batch, tableName!, columnName!, criteria!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ApplyFilterValues(TableCommands commands, string sessionId, string? tableName, string? columnName, string? filterValuesJson)
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
            throw new ArgumentException($"Invalid JSON array for filterValues: {ex.Message}", nameof(filterValuesJson));
        }

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.ApplyFilter(batch, tableName!, columnName!, filterValues));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ClearFilters(TableCommands commands, string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "clear-filters");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.ClearFilters(batch, tableName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string GetFilters(TableCommands commands, string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "get-filters");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetFilters(batch, tableName!));

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

    private static string AddColumn(TableCommands commands, string sessionId, string? tableName, string? columnName, string? positionStr)
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
                throw new ArgumentException($"Invalid position value: '{positionStr}'. Must be a number.", nameof(positionStr));
            }
        }

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.AddColumn(batch, tableName!, columnName!, position));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string RemoveColumn(TableCommands commands, string sessionId, string? tableName, string? columnName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "remove-column");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "remove-column");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.RemoveColumn(batch, tableName!, columnName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string RenameColumn(TableCommands commands, string sessionId, string? tableName, string? oldColumnName, string? newColumnName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "rename-column");
        if (string.IsNullOrWhiteSpace(oldColumnName)) ExcelToolsBase.ThrowMissingParameter(nameof(oldColumnName), "rename-column");
        if (string.IsNullOrWhiteSpace(newColumnName)) ExcelToolsBase.ThrowMissingParameter(nameof(newColumnName), "rename-column");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.RenameColumn(batch, tableName!, oldColumnName!, newColumnName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === PHASE 2: STRUCTURED REFERENCE & SORT OPERATIONS ===

    private static string GetStructuredReference(TableCommands commands, string sessionId, string? tableName, string? regionStr, string? columnName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "get-structured-reference");

        // Parse region string to enum (default: Data)
        var region = Core.Models.TableRegion.Data; // Default
        if (!string.IsNullOrWhiteSpace(regionStr) && !Enum.TryParse(regionStr, true, out region))
        {
            throw new ArgumentException($"Invalid region '{regionStr}'. Valid values: All, Data, Headers, Totals, ThisRow", nameof(regionStr));
        }

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetStructuredReference(batch, tableName!, region, columnName));

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

    private static string SortTable(TableCommands commands, string sessionId, string? tableName, string? columnName, bool ascending)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "sort");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "sort");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Sort(batch, tableName!, columnName!, ascending));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string SortTableMulti(TableCommands commands, string sessionId, string? tableName, string? sortColumnsJson)
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
                throw new ArgumentException("sortColumns JSON must be a non-empty array", nameof(sortColumnsJson));
            }
        }
        catch (JsonException ex)
        {
            throw new ArgumentException($"Invalid sortColumns JSON: {ex.Message}", nameof(sortColumnsJson));
        }

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Sort(batch, tableName!, sortColumns));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === NUMBER FORMAT OPERATIONS ===

    private static string GetColumnNumberFormat(TableCommands commands, string sessionId, string? tableName, string? columnName)
    {
        if (string.IsNullOrEmpty(tableName))
            ExcelToolsBase.ThrowMissingParameter("tableName", "get-column-number-format");
        if (string.IsNullOrEmpty(columnName))
            ExcelToolsBase.ThrowMissingParameter("columnName", "get-column-number-format");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetColumnNumberFormat(batch, tableName!, columnName!));

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

    private static string SetColumnNumberFormat(TableCommands commands, string sessionId, string? tableName, string? columnName, string? formatCode)
    {
        if (string.IsNullOrEmpty(tableName))
            ExcelToolsBase.ThrowMissingParameter("tableName", "set-column-number-format");
        if (string.IsNullOrEmpty(columnName))
            ExcelToolsBase.ThrowMissingParameter("columnName", "set-column-number-format");
        if (string.IsNullOrEmpty(formatCode))
            ExcelToolsBase.ThrowMissingParameter("formatCode", "set-column-number-format");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.SetColumnNumberFormat(batch, tableName!, columnName!, formatCode!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }
}

