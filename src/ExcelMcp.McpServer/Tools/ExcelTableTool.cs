using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel Table lifecycle and data operations.
/// </summary>
[McpServerToolType]
public static partial class TableTool
{
    /// <summary>
    /// Excel Tables (ListObjects) - lifecycle and data operations.
    ///
    /// BEST PRACTICE: Use 'list' to check existing tables before creating.
    /// Prefer 'append'/'resize'/'rename' over delete+recreate to preserve references.
    ///
    /// WARNING: Deleting tables used as PivotTable sources or in Data Model relationships will break those objects.
    ///
    /// CREATING TABLES: Specify sheetName, tableName, and rangeAddress. Set hasHeaders=true if first row contains headers.
    ///
    /// DATA MODEL WORKFLOW: To analyze worksheet data with DAX/Power Pivot:
    /// 1. Create or identify an Excel Table on a worksheet
    /// 2. Use 'add-to-datamodel' action to add the table to Power Pivot
    /// 3. Then use excel_datamodel to create DAX measures on it
    ///
    /// DAX-BACKED TABLES: Create tables populated by DAX EVALUATE queries:
    /// - 'create-from-dax': Create a new table backed by a DAX query (e.g., SUMMARIZE, FILTER)
    /// - 'update-dax': Update the DAX query for an existing DAX-backed table
    /// - 'get-dax': Get the DAX query info for a table (check if it's DAX-backed)
    ///
    /// APPENDING DATA: Use 'append' action with csvData in simple CSV format (comma-separated, newline-separated rows).
    ///
    /// Related: excel_table_column (filter/sort/columns), excel_datamodel (DAX measures, evaluate queries)
    /// </summary>
    /// <param name="action">The table operation to perform</param>
    /// <param name="excelPath">Full path to Excel file (for reference/logging)</param>
    /// <param name="sessionId">Session ID from excel_file 'open'. Required for all actions.</param>
    /// <param name="tableName">Name of the table to operate on. Required for: read, rename, delete, resize, toggle-totals, set-column-total, append, get-data, set-style, add-to-datamodel, update-dax, get-dax. Used as new table name for: create, create-from-dax</param>
    /// <param name="sheetName">Name of the worksheet containing the table. Required for: create, create-from-dax</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1:D10'). Required for: create, resize. Optional targetCell for: create-from-dax (default: 'A1')</param>
    /// <param name="newName">New name for the table. Required for: rename. Column name for: set-column-total</param>
    /// <param name="hasHeaders">For 'create': true if first row contains column headers (default: true). For 'toggle-totals': true to show totals row, false to hide.</param>
    /// <param name="styleName">Multi-purpose string parameter: Table style name for 'create'/'set-style' (e.g., 'TableStyleMedium2'). Total function for 'set-column-total' (Sum, Average, Count, etc.). CSV data for 'append'.</param>
    /// <param name="visibleOnly">For 'get-data': true to return only visible (non-filtered) rows. Default: false (all rows)</param>
    /// <param name="daxQuery">DAX EVALUATE query for 'create-from-dax' and 'update-dax' actions (e.g., 'EVALUATE SUMMARIZE(...))'</param>
    [McpServerTool(Name = "excel_table", Title = "Excel Table Operations")]
    [McpMeta("category", "data")]
    [McpMeta("requiresSession", true)]
    public static partial string Table(
        TableAction action,
        string excelPath,
        string sessionId,
        [DefaultValue(null)] string? tableName,
        [DefaultValue(null)] string? sheetName,
        [DefaultValue(null)] string? rangeAddress,
        [DefaultValue(null)] string? newName,
        [DefaultValue(true)] bool hasHeaders,
        [DefaultValue(null)] string? styleName,
        [DefaultValue(false)] bool visibleOnly,
        [DefaultValue(null)] string? daxQuery)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_table",
            action.ToActionString(),
            excelPath,
            () =>
            {
                var tableCommands = new TableCommands();

                return action switch
                {
                    TableAction.List => ListTables(tableCommands, sessionId),
                    TableAction.Create => CreateTable(tableCommands, sessionId, sheetName, tableName, rangeAddress, hasHeaders, styleName),
                    TableAction.Read => ReadTable(tableCommands, sessionId, tableName),
                    TableAction.Rename => RenameTable(tableCommands, sessionId, tableName, newName),
                    TableAction.Delete => DeleteTable(tableCommands, sessionId, tableName),
                    TableAction.Resize => ResizeTable(tableCommands, sessionId, tableName, rangeAddress),
                    TableAction.ToggleTotals => ToggleTotals(tableCommands, sessionId, tableName, hasHeaders),
                    TableAction.SetColumnTotal => SetColumnTotal(tableCommands, sessionId, tableName, newName, styleName),
                    TableAction.Append => AppendRows(tableCommands, sessionId, tableName, styleName),
                    TableAction.GetData => GetData(tableCommands, sessionId, tableName, visibleOnly),
                    TableAction.SetStyle => SetTableStyle(tableCommands, sessionId, tableName, styleName),
                    TableAction.AddToDataModel => AddToDataModel(tableCommands, sessionId, tableName),
                    TableAction.CreateFromDax => CreateFromDaxAction(tableCommands, sessionId, sheetName, tableName, daxQuery, rangeAddress),
                    TableAction.UpdateDax => UpdateDaxAction(tableCommands, sessionId, tableName, daxQuery),
                    TableAction.GetDax => GetDaxAction(tableCommands, sessionId, tableName),
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

    private static string CreateTable(TableCommands commands, string sessionId, string? sheetName, string? tableName, string? rangeAddress, bool hasHeaders, string? styleName)
    {
        if (string.IsNullOrWhiteSpace(sheetName)) ExcelToolsBase.ThrowMissingParameter(nameof(sheetName), "create");
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "create");
        if (string.IsNullOrWhiteSpace(rangeAddress)) ExcelToolsBase.ThrowMissingParameter(nameof(rangeAddress), "create");

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.Create(batch, sheetName!, tableName!, rangeAddress!, hasHeaders, styleName);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Table created successfully." }, ExcelToolsBase.JsonOptions);
        }
#pragma warning disable CA1031 // MCP protocol requires JSON error responses, not thrown exceptions
        catch (Exception ex)
#pragma warning restore CA1031
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string ReadTable(TableCommands commands, string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "read");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Read(batch, tableName!));

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string GetData(TableCommands commands, string sessionId, string? tableName, bool visibleOnly)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "get-data");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetData(batch, tableName!, visibleOnly));

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Table renamed successfully." }, ExcelToolsBase.JsonOptions);
        }
#pragma warning disable CA1031 // MCP protocol requires JSON error responses, not thrown exceptions
        catch (Exception ex)
#pragma warning restore CA1031
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
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

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Table deleted successfully." }, ExcelToolsBase.JsonOptions);
        }
#pragma warning disable CA1031 // MCP protocol requires JSON error responses, not thrown exceptions
        catch (Exception ex)
#pragma warning restore CA1031
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string ResizeTable(TableCommands commands, string sessionId, string? tableName, string? rangeAddress)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "resize");
        if (string.IsNullOrWhiteSpace(rangeAddress)) ExcelToolsBase.ThrowMissingParameter(nameof(rangeAddress), "resize");

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.Resize(batch, tableName!, rangeAddress!);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Table resized successfully." }, ExcelToolsBase.JsonOptions);
        }
#pragma warning disable CA1031 // MCP protocol requires JSON error responses, not thrown exceptions
        catch (Exception ex)
#pragma warning restore CA1031
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
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

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Totals toggled successfully." }, ExcelToolsBase.JsonOptions);
        }
#pragma warning disable CA1031 // MCP protocol requires JSON error responses, not thrown exceptions
        catch (Exception ex)
#pragma warning restore CA1031
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
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

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Column total set successfully." }, ExcelToolsBase.JsonOptions);
        }
#pragma warning disable CA1031 // MCP protocol requires JSON error responses, not thrown exceptions
        catch (Exception ex)
#pragma warning restore CA1031
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
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

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Rows appended successfully." }, ExcelToolsBase.JsonOptions);
        }
#pragma warning disable CA1031 // MCP protocol requires JSON error responses, not thrown exceptions
        catch (Exception ex)
#pragma warning restore CA1031
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
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

    private static string SetTableStyle(TableCommands commands, string sessionId, string? tableName, string? styleName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "set-style");
        if (string.IsNullOrWhiteSpace(styleName)) ExcelToolsBase.ThrowMissingParameter(nameof(styleName), "set-style");

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.SetStyle(batch, tableName!, styleName!);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Table style set successfully." }, ExcelToolsBase.JsonOptions);
        }
#pragma warning disable CA1031 // MCP protocol requires JSON error responses, not thrown exceptions
        catch (Exception ex)
#pragma warning restore CA1031
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
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

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = "Table added to data model successfully." }, ExcelToolsBase.JsonOptions);
        }
#pragma warning disable CA1031 // MCP protocol requires JSON error responses, not thrown exceptions
        catch (Exception ex)
#pragma warning restore CA1031
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string CreateFromDaxAction(TableCommands commands, string sessionId, string? sheetName, string? tableName, string? daxQuery, string? targetCell)
    {
        if (string.IsNullOrWhiteSpace(sheetName)) ExcelToolsBase.ThrowMissingParameter(nameof(sheetName), "create-from-dax");
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "create-from-dax");
        if (string.IsNullOrWhiteSpace(daxQuery)) ExcelToolsBase.ThrowMissingParameter(nameof(daxQuery), "create-from-dax");

        // Use targetCell if provided, otherwise default to "A1"
        var cellAddress = string.IsNullOrWhiteSpace(targetCell) ? "A1" : targetCell;

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.CreateFromDax(batch, sheetName!, tableName!, daxQuery!, cellAddress);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = $"DAX-backed table '{tableName}' created successfully." }, ExcelToolsBase.JsonOptions);
        }
#pragma warning disable CA1031 // MCP protocol requires JSON error responses, not thrown exceptions
        catch (Exception ex)
#pragma warning restore CA1031
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string UpdateDaxAction(TableCommands commands, string sessionId, string? tableName, string? daxQuery)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "update-dax");
        if (string.IsNullOrWhiteSpace(daxQuery)) ExcelToolsBase.ThrowMissingParameter(nameof(daxQuery), "update-dax");

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    commands.UpdateDax(batch, tableName!, daxQuery!);
                    return 0;
                });

            return JsonSerializer.Serialize(new OperationResult { Success = true, Message = $"DAX query updated for table '{tableName}'." }, ExcelToolsBase.JsonOptions);
        }
#pragma warning disable CA1031 // MCP protocol requires JSON error responses, not thrown exceptions
        catch (Exception ex)
#pragma warning restore CA1031
        {
            return JsonSerializer.Serialize(new OperationResult { Success = false, ErrorMessage = ex.Message }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string GetDaxAction(TableCommands commands, string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "get-dax");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetDax(batch, tableName!));

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}

