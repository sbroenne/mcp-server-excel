using System.ComponentModel;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel range edit operations - insert, delete, find, replace, sort.
/// </summary>
[McpServerToolType]
public static partial class ExcelRangeEditTool
{
    /// <summary>
    /// Range editing operations: insert/delete cells, rows, and columns; find/replace text; sort data.
    ///
    /// INSERT/DELETE CELLS: Specify shiftDirection to control how surrounding cells move.
    /// - Insert: 'Down' or 'Right'
    /// - Delete: 'Up' or 'Left'
    ///
    /// INSERT/DELETE ROWS: Use row range like '5:10' to insert/delete rows 5-10.
    /// INSERT/DELETE COLUMNS: Use column range like 'B:D' to insert/delete columns B-D.
    ///
    /// FIND/REPLACE: Search within the specified range with optional case/cell matching.
    /// - Find returns up to 10 matching cell addresses with total count.
    /// - Replace modifies all matches by default (replaceAll=true).
    ///
    /// SORT: Specify sortColumns as array of {columnIndex: 1, ascending: true} objects.
    /// Column indices are 1-based relative to the range.
    /// </summary>
    /// <param name="action">The range edit operation to perform</param>
    /// <param name="path">Full path to the Excel workbook file (e.g., 'C:\Reports\Sales.xlsx')</param>
    /// <param name="sessionId">Session identifier returned from excel_file open action - required for all operations</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1:D10', 'B:D' for columns, '5:10' for rows)</param>
    /// <param name="shiftDirection">Direction to shift cells: 'Down' or 'Right' for insert, 'Up' or 'Left' for delete</param>
    /// <param name="searchValue">Text or value to search for in find/replace operations</param>
    /// <param name="replaceValue">Text or value to replace matches with in replace operation</param>
    /// <param name="matchCase">Whether to match case exactly (default: false = case-insensitive)</param>
    /// <param name="matchEntireCell">Whether to match entire cell content only (default: false = partial match)</param>
    /// <param name="searchInFormulas">Whether to search within formula text (default: true)</param>
    /// <param name="searchInValues">Whether to search within displayed values (default: true)</param>
    /// <param name="replaceAll">Whether to replace all occurrences or just the first one (default: true)</param>
    /// <param name="sortColumns">Array of sort specifications: [{columnIndex: 1, ascending: true}, ...] - columnIndex is 1-based relative to range</param>
    /// <param name="hasHeaderRow">Whether the range has a header row to exclude from sorting (default: true)</param>
    [McpServerTool(Name = "excel_range_edit", Title = "Excel Range Edit Operations", Destructive = true)]
    [McpMeta("category", "data")]
    [McpMeta("requiresSession", true)]
    public static partial string RangeEdit(
        RangeEditAction action,
        string path,
        string sessionId,
        [DefaultValue(null)] string? sheetName,
        [DefaultValue(null)] string? rangeAddress,
        [DefaultValue(null)] string? shiftDirection,
        [DefaultValue(null)] string? searchValue,
        [DefaultValue(null)] string? replaceValue,
        [DefaultValue(null)] bool? matchCase,
        [DefaultValue(null)] bool? matchEntireCell,
        [DefaultValue(null)] bool? searchInFormulas,
        [DefaultValue(null)] bool? searchInValues,
        [DefaultValue(null)] bool? replaceAll,
        [DefaultValue(null)] List<SortColumn>? sortColumns,
        [DefaultValue(null)] bool? hasHeaderRow)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_range_edit",
            ServiceRegistry.RangeEdit.ToActionString(action),
            path,
            () => action switch
            {
                RangeEditAction.InsertCells => ForwardInsertCells(sessionId, sheetName, rangeAddress, shiftDirection),
                RangeEditAction.DeleteCells => ForwardDeleteCells(sessionId, sheetName, rangeAddress, shiftDirection),
                RangeEditAction.InsertRows => ForwardSimpleRange(sessionId, sheetName, rangeAddress, "insert-rows"),
                RangeEditAction.DeleteRows => ForwardSimpleRange(sessionId, sheetName, rangeAddress, "delete-rows"),
                RangeEditAction.InsertColumns => ForwardSimpleRange(sessionId, sheetName, rangeAddress, "insert-columns"),
                RangeEditAction.DeleteColumns => ForwardSimpleRange(sessionId, sheetName, rangeAddress, "delete-columns"),
                RangeEditAction.Find => ForwardFind(sessionId, sheetName, rangeAddress, searchValue, matchCase, matchEntireCell, searchInFormulas, searchInValues),
                RangeEditAction.Replace => ForwardReplace(sessionId, sheetName, rangeAddress, searchValue, replaceValue, matchCase, matchEntireCell, replaceAll),
                RangeEditAction.Sort => ForwardSort(sessionId, sheetName, rangeAddress, sortColumns, hasHeaderRow),
                _ => throw new ArgumentException($"Unknown action: {action} ({ServiceRegistry.RangeEdit.ToActionString(action)})", nameof(action))
            });
    }

    private static string ForwardInsertCells(string sessionId, string? sheetName, string? rangeAddress, string? shiftDirection)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "insert-cells");
        if (string.IsNullOrEmpty(shiftDirection))
            ExcelToolsBase.ThrowMissingParameter("shiftDirection", "insert-cells");

        return ExcelToolsBase.ForwardToService("range.insert-cells", sessionId, new
        {
            sheetName = sheetName ?? "",
            range = rangeAddress,
            shiftDirection
        });
    }

    private static string ForwardDeleteCells(string sessionId, string? sheetName, string? rangeAddress, string? shiftDirection)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "delete-cells");
        if (string.IsNullOrEmpty(shiftDirection))
            ExcelToolsBase.ThrowMissingParameter("shiftDirection", "delete-cells");

        return ExcelToolsBase.ForwardToService("range.delete-cells", sessionId, new
        {
            sheetName = sheetName ?? "",
            range = rangeAddress,
            shiftDirection
        });
    }

    private static string ForwardSimpleRange(string sessionId, string? sheetName, string? rangeAddress, string actionName)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", actionName);

        return ExcelToolsBase.ForwardToService($"range.{actionName}", sessionId, new
        {
            sheetName = sheetName ?? "",
            range = rangeAddress
        });
    }

    private static string ForwardFind(string sessionId, string? sheetName, string? rangeAddress, string? searchValue,
        bool? matchCase, bool? matchEntireCell, bool? searchInFormulas, bool? searchInValues)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "find");
        if (string.IsNullOrEmpty(searchValue))
            ExcelToolsBase.ThrowMissingParameter("searchValue", "find");

        return ExcelToolsBase.ForwardToService("range.find", sessionId, new
        {
            sheetName = sheetName ?? "",
            range = rangeAddress,
            searchValue,
            matchCase = matchCase ?? false,
            matchEntireCell = matchEntireCell ?? false,
            searchFormulas = searchInFormulas ?? true,
            searchValues = searchInValues ?? true
        });
    }

    private static string ForwardReplace(string sessionId, string? sheetName, string? rangeAddress,
        string? searchValue, string? replaceValue, bool? matchCase, bool? matchEntireCell, bool? replaceAll)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "replace");
        if (string.IsNullOrEmpty(searchValue))
            ExcelToolsBase.ThrowMissingParameter("searchValue", "replace");
        if (replaceValue == null)
            ExcelToolsBase.ThrowMissingParameter("replaceValue", "replace");

        return ExcelToolsBase.ForwardToService("range.replace", sessionId, new
        {
            sheetName = sheetName ?? "",
            range = rangeAddress,
            findValue = searchValue,
            replaceValue,
            matchCase = matchCase ?? false,
            matchEntireCell = matchEntireCell ?? false,
            replaceAll = replaceAll ?? true
        });
    }

    private static string ForwardSort(string sessionId, string? sheetName, string? rangeAddress,
        List<SortColumn>? sortColumns, bool? hasHeaderRow)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "sort");
        if (sortColumns == null || sortColumns.Count == 0)
            ExcelToolsBase.ThrowMissingParameter("sortColumns", "sort");

        return ExcelToolsBase.ForwardToService("range.sort", sessionId, new
        {
            sheetName = sheetName ?? "",
            range = rangeAddress,
            sortColumns = sortColumns!.Select(sc => new { columnIndex = sc.ColumnIndex, ascending = sc.Ascending }).ToList(),
            hasHeaders = hasHeaderRow ?? true
        });
    }
}




