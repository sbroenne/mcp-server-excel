using System.ComponentModel;
using ModelContextProtocol.Server;

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
    /// <param name="path">Full path to Excel file (for reference/logging)</param>
    /// <param name="sessionId">Session ID from excel_file 'open'. Required for all actions.</param>
    /// <param name="tableName">Name of the table to operate on. Required for: read, rename, delete, resize, toggle-totals, set-column-total, append, get-data, set-style, add-to-datamodel, update-dax, get-dax. Used as new table name for: create, create-from-dax</param>
    /// <param name="sheetName">Name of the worksheet containing the table. Required for: create, create-from-dax</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1:D10'). Required for: create, resize. Optional targetCell for: create-from-dax (default: 'A1')</param>
    /// <param name="newName">New name for the table. Required for: rename. Column name for: set-column-total</param>
    /// <param name="hasHeaders">For 'create': true if first row contains column headers (default: true). For 'toggle-totals': true to show totals row, false to hide.</param>
    /// <param name="styleName">Table style name for 'create'/'set-style' (e.g., 'TableStyleMedium2'). Total function for 'set-column-total' (Sum, Average, Count, etc.).</param>
    /// <param name="visibleOnly">For 'get-data': true to return only visible (non-filtered) rows. Default: false (all rows)</param>
    /// <param name="daxQuery">DAX EVALUATE query for 'create-from-dax' and 'update-dax' actions (e.g., 'EVALUATE SUMMARIZE(...))'</param>
    /// <param name="csvData">CSV data for 'append' action. Simple format: comma-separated values, newline-separated rows. First row should match table column order.</param>
    [McpServerTool(Name = "excel_table", Title = "Excel Table Operations", Destructive = true)]
    [McpMeta("category", "data")]
    [McpMeta("requiresSession", true)]
    public static partial string Table(
        TableAction action,
        string path,
        string sessionId,
        [DefaultValue(null)] string? tableName,
        [DefaultValue(null)] string? sheetName,
        [DefaultValue(null)] string? rangeAddress,
        [DefaultValue(null)] string? newName,
        [DefaultValue(true)] bool hasHeaders,
        [DefaultValue(null)] string? styleName,
        [DefaultValue(false)] bool visibleOnly,
        [DefaultValue(null)] string? daxQuery,
        [DefaultValue(null)] string? csvData)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_table",
            ServiceRegistry.Table.ToActionString(action),
            path,
            () =>
            {
                return action switch
                {
                    TableAction.List => ForwardList(sessionId),
                    TableAction.Create => ForwardCreate(sessionId, sheetName, tableName, rangeAddress, hasHeaders, styleName),
                    TableAction.Read => ForwardRead(sessionId, tableName),
                    TableAction.Rename => ForwardRename(sessionId, tableName, newName),
                    TableAction.Delete => ForwardDelete(sessionId, tableName),
                    TableAction.Resize => ForwardResize(sessionId, tableName, rangeAddress),
                    TableAction.ToggleTotals => ForwardToggleTotals(sessionId, tableName, hasHeaders),
                    TableAction.SetColumnTotal => ForwardSetColumnTotal(sessionId, tableName, newName, styleName),
                    TableAction.Append => ForwardAppend(sessionId, tableName, csvData),
                    TableAction.GetData => ForwardGetData(sessionId, tableName, visibleOnly),
                    TableAction.SetStyle => ForwardSetStyle(sessionId, tableName, styleName),
                    TableAction.AddToDataModel => ForwardAddToDataModel(sessionId, tableName),
                    TableAction.CreateFromDax => ForwardCreateFromDax(sessionId, sheetName, tableName, daxQuery, rangeAddress),
                    TableAction.UpdateDax => ForwardUpdateDax(sessionId, tableName, daxQuery),
                    TableAction.GetDax => ForwardGetDax(sessionId, tableName),
                    _ => throw new ArgumentException(
                        $"Unknown action: {action} ({ServiceRegistry.Table.ToActionString(action)})", nameof(action))
                };
            });
    }

    // === SERVICE FORWARDING METHODS ===

    private static string ForwardList(string sessionId)
    {
        return ExcelToolsBase.ForwardToService("table.list", sessionId);
    }

    private static string ForwardCreate(string sessionId, string? sheetName, string? tableName, string? rangeAddress, bool hasHeaders, string? styleName)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            ExcelToolsBase.ThrowMissingParameter(nameof(sheetName), "create");
        if (string.IsNullOrWhiteSpace(tableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "create");
        if (string.IsNullOrWhiteSpace(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter(nameof(rangeAddress), "create");

        return ExcelToolsBase.ForwardToService("table.create", sessionId, new
        {
            sheetName,
            tableName,
            range = rangeAddress,
            hasHeaders,
            tableStyle = styleName
        });
    }

    private static string ForwardRead(string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "read");

        return ExcelToolsBase.ForwardToService("table.read", sessionId, new { tableName });
    }

    private static string ForwardRename(string sessionId, string? tableName, string? newName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "rename");
        if (string.IsNullOrWhiteSpace(newName))
            ExcelToolsBase.ThrowMissingParameter(nameof(newName), "rename");

        return ExcelToolsBase.ForwardToService("table.rename", sessionId, new { tableName, newName });
    }

    private static string ForwardDelete(string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "delete");

        return ExcelToolsBase.ForwardToService("table.delete", sessionId, new { tableName });
    }

    private static string ForwardResize(string sessionId, string? tableName, string? rangeAddress)
    {
        if (string.IsNullOrWhiteSpace(tableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "resize");
        if (string.IsNullOrWhiteSpace(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter(nameof(rangeAddress), "resize");

        return ExcelToolsBase.ForwardToService("table.resize", sessionId, new { tableName, newRange = rangeAddress });
    }

    private static string ForwardToggleTotals(string sessionId, string? tableName, bool showTotals)
    {
        if (string.IsNullOrWhiteSpace(tableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "toggle-totals");

        return ExcelToolsBase.ForwardToService("table.toggle-totals", sessionId, new { tableName, showTotals });
    }

    private static string ForwardSetColumnTotal(string sessionId, string? tableName, string? columnName, string? totalFunction)
    {
        if (string.IsNullOrWhiteSpace(tableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "set-column-total");
        if (string.IsNullOrWhiteSpace(columnName))
            ExcelToolsBase.ThrowMissingParameter("columnName (newName parameter)", "set-column-total");
        if (string.IsNullOrWhiteSpace(totalFunction))
            ExcelToolsBase.ThrowMissingParameter("totalFunction (styleName parameter)", "set-column-total");

        return ExcelToolsBase.ForwardToService("table.set-column-total", sessionId, new { tableName, columnName, totalFunction });
    }

    private static string ForwardAppend(string sessionId, string? tableName, string? csvData)
    {
        if (string.IsNullOrWhiteSpace(tableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "append");
        if (string.IsNullOrWhiteSpace(csvData))
            ExcelToolsBase.ThrowMissingParameter(nameof(csvData), "append");

        // Parse CSV data to List<List<object?>>
        var rows = ParseCsvToRows(csvData!);

        return ExcelToolsBase.ForwardToService("table.append", sessionId, new { tableName, rows });
    }

    private static string ForwardGetData(string sessionId, string? tableName, bool visibleOnly)
    {
        if (string.IsNullOrWhiteSpace(tableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "get-data");

        return ExcelToolsBase.ForwardToService("table.get-data", sessionId, new { tableName, visibleOnly });
    }

    private static string ForwardSetStyle(string sessionId, string? tableName, string? styleName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "set-style");
        if (string.IsNullOrWhiteSpace(styleName))
            ExcelToolsBase.ThrowMissingParameter(nameof(styleName), "set-style");

        return ExcelToolsBase.ForwardToService("table.set-style", sessionId, new { tableName, tableStyle = styleName });
    }

    private static string ForwardAddToDataModel(string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "add-to-datamodel");

        return ExcelToolsBase.ForwardToService("table.add-to-datamodel", sessionId, new { tableName });
    }

    private static string ForwardCreateFromDax(string sessionId, string? sheetName, string? tableName, string? daxQuery, string? targetCell)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            ExcelToolsBase.ThrowMissingParameter(nameof(sheetName), "create-from-dax");
        if (string.IsNullOrWhiteSpace(tableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "create-from-dax");
        if (string.IsNullOrWhiteSpace(daxQuery))
            ExcelToolsBase.ThrowMissingParameter(nameof(daxQuery), "create-from-dax");

        // Use targetCell if provided, otherwise default to "A1"
        var cellAddress = string.IsNullOrWhiteSpace(targetCell) ? "A1" : targetCell;

        return ExcelToolsBase.ForwardToService("table.create-from-dax", sessionId, new { sheetName, tableName, daxQuery, targetCell = cellAddress });
    }

    private static string ForwardUpdateDax(string sessionId, string? tableName, string? daxQuery)
    {
        if (string.IsNullOrWhiteSpace(tableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "update-dax");
        if (string.IsNullOrWhiteSpace(daxQuery))
            ExcelToolsBase.ThrowMissingParameter(nameof(daxQuery), "update-dax");

        return ExcelToolsBase.ForwardToService("table.update-dax", sessionId, new { tableName, daxQuery });
    }

    private static string ForwardGetDax(string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "get-dax");

        return ExcelToolsBase.ForwardToService("table.get-dax", sessionId, new { tableName });
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
}





