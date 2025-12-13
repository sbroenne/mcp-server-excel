using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands.Table;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel Table (ListObject) operations - structured data with AutoFilter and dynamic expansion.
/// </summary>
[McpServerToolType]
public static partial class TableTool
{
    /// <summary>
    /// Manage Excel Tables (ListObjects) - structured data with AutoFilter.
    /// DATA ACCESS: Use action 'get-data' to return table rows. Set visibleOnly=true to respect active filters.
    /// </summary>
    /// <param name="action">Action to perform</param>
    /// <param name="excelPath">Excel file path (.xlsx or .xlsm)</param>
    /// <param name="sessionId">Session ID from excel_file 'open' action</param>
    /// <param name="tableName">Table name (required for most actions). Must start with letter/underscore, alphanumeric + underscore only</param>
    /// <param name="sheetName">Sheet name (required for create)</param>
    /// <param name="range">Excel range (e.g., 'A1:D10') - required for create/resize</param>
    /// <param name="newName">New table name (required for rename) or column name (required for add-column, rename-column, etc.). Table names must follow Excel naming rules (start with letter/underscore, alphanumeric only). Column names can be any string including numbers.</param>
    /// <param name="hasHeaders">Whether the range has headers (default: true for create) or show totals (for toggle-totals)</param>
    /// <param name="tableStyle">Table style name (e.g., 'TableStyleMedium2') for create/set-style, or total function (sum/avg/count) for set-column-total, or CSV data for append</param>
    /// <param name="filterCriteria">Filter criteria (e.g., '>100', '=Text') for apply-filter, or column position (0-based) for add-column</param>
    /// <param name="filterValues">JSON array of filter values (e.g., '["Value1","Value2"]') for apply-filter-values</param>
    /// <param name="formatCode">Excel format code for set-column-number-format. ALWAYS use US format codes (e.g., '$#,##0.00', '0.00%', 'm/d/yyyy'). The server auto-translates to the user's locale.</param>
    /// <param name="visibleOnly">When reading data, return only rows currently visible after filters (default: false)</param>
    [McpServerTool(Name = "excel_table", Title = "Excel Table Operations")]
    [McpMeta("category", "data")]
    [McpMeta("requiresSession", true)]
    public static partial string Table(
        TableAction action,
        string excelPath,
        string sessionId,
        [DefaultValue(null)] string? tableName,
        [DefaultValue(null)] string? sheetName,
        [DefaultValue(null)] string? range,
        [DefaultValue(null)] string? newName,
        [DefaultValue(true)] bool hasHeaders,
        [DefaultValue(null)] string? tableStyle,
        [DefaultValue(null)] string? filterCriteria,
        [DefaultValue(null)] string? filterValues,
        [DefaultValue(null)] string? formatCode,
        [DefaultValue(false)] bool visibleOnly)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_table",
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
                    TableAction.GetData => GetData(tableCommands, sessionId, tableName, visibleOnly),
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

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.Create(batch, sheetName!, tableName!, range!, hasHeaders, tableStyle);
                    return 0;
                });

            return JsonSerializer.Serialize(new { success = true, message = "Table created successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, errorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
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

    private static string GetData(TableCommands commands, string sessionId, string? tableName, bool visibleOnly)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "get-data");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetData(batch, tableName!, visibleOnly));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.TableName,
            result.Headers,
            result.Data,
            result.RowCount,
            result.ColumnCount,
            VisibleOnly = visibleOnly,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string RenameTable(TableCommands commands, string sessionId, string? tableName, string? newName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "rename");
        if (string.IsNullOrWhiteSpace(newName)) ExcelToolsBase.ThrowMissingParameter(nameof(newName), "rename");

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.Rename(batch, tableName!, newName!);
                    return 0;
                });

            return JsonSerializer.Serialize(new { success = true, message = "Table renamed successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, errorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string DeleteTable(TableCommands commands, string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "delete");

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.Delete(batch, tableName!);
                    return 0;
                });

            return JsonSerializer.Serialize(new { success = true, message = "Table deleted successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, errorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string ResizeTable(TableCommands commands, string sessionId, string? tableName, string? newRange)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "resize");
        if (string.IsNullOrWhiteSpace(newRange)) ExcelToolsBase.ThrowMissingParameter(nameof(newRange), "resize");

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.Resize(batch, tableName!, newRange!);
                    return 0;
                });

            return JsonSerializer.Serialize(new { success = true, message = "Table resized successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, errorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string ToggleTotals(TableCommands commands, string sessionId, string? tableName, bool showTotals)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "toggle-totals");

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.ToggleTotals(batch, tableName!, showTotals);
                    return 0;
                });

            return JsonSerializer.Serialize(new { success = true, message = "Totals toggled successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, errorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string SetColumnTotal(TableCommands commands, string sessionId, string? tableName, string? columnName, string? totalFunction)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "set-column-total");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "set-column-total");
        if (string.IsNullOrWhiteSpace(totalFunction)) ExcelToolsBase.ThrowMissingParameter(nameof(totalFunction), "set-column-total");

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.SetColumnTotal(batch, tableName!, columnName!, totalFunction!);
                    return 0;
                });

            return JsonSerializer.Serialize(new { success = true, message = "Column total set successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, errorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string AppendRows(TableCommands commands, string sessionId, string? tableName, string? csvData)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "append");
        if (string.IsNullOrWhiteSpace(csvData)) ExcelToolsBase.ThrowMissingParameter(nameof(csvData), "append");

        try
        {
            // Parse CSV data to List<List<object?>>
            var rows = ParseCsvToRows(csvData!);

            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.Append(batch, tableName!, rows);
                    return 0;
                });

            return JsonSerializer.Serialize(new { success = true, message = "Rows appended successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, errorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
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

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.SetStyle(batch, tableName!, tableStyle!);
                    return 0;
                });

            return JsonSerializer.Serialize(new { success = true, message = "Table style set successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, errorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string AddToDataModel(TableCommands commands, string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "add-to-datamodel");

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.AddToDataModel(batch, tableName!);
                    return 0;
                });

            return JsonSerializer.Serialize(new { success = true, message = "Table added to data model successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, errorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    // === FILTER OPERATIONS ===

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

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = "Filter applied successfully."
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message
            }, ExcelToolsBase.JsonOptions);
        }
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

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.ApplyFilter(batch, tableName!, columnName!, filterValues);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = "Filter applied successfully."
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message
            }, ExcelToolsBase.JsonOptions);
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

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = "Filters cleared successfully."
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message
            }, ExcelToolsBase.JsonOptions);
        }
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

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.AddColumn(batch, tableName!, columnName!, position);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = "Column added successfully."
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message
            }, ExcelToolsBase.JsonOptions);
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

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = "Column removed successfully."
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message
            }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string RenameColumn(TableCommands commands, string sessionId, string? tableName, string? oldColumnName, string? newColumnName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "rename-column");
        if (string.IsNullOrWhiteSpace(oldColumnName)) ExcelToolsBase.ThrowMissingParameter(nameof(oldColumnName), "rename-column");
        if (string.IsNullOrWhiteSpace(newColumnName)) ExcelToolsBase.ThrowMissingParameter(nameof(newColumnName), "rename-column");

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.RenameColumn(batch, tableName!, oldColumnName!, newColumnName!);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = "Column renamed successfully."
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message
            }, ExcelToolsBase.JsonOptions);
        }
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

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.Sort(batch, tableName!, columnName!, ascending);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = "Table sorted successfully."
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message
            }, ExcelToolsBase.JsonOptions);
        }
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

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.Sort(batch, tableName!, sortColumns);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = "Table sorted successfully."
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message
            }, ExcelToolsBase.JsonOptions);
        }
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

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.SetColumnNumberFormat(batch, tableName!, columnName!, formatCode!);
                    return 0;
                });

            return JsonSerializer.Serialize(new { success = true, message = "Column number format set successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, errorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }
}

