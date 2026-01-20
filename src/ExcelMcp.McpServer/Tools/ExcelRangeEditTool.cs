using System.ComponentModel;
using System.Text.Json;
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
    /// <param name="excelPath">Full path to the Excel workbook file (e.g., 'C:\Reports\Sales.xlsx')</param>
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
        string excelPath,
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
            action.ToActionString(),
            excelPath,
            () =>
            {
                var rangeCommands = new RangeCommands();

                return action switch
                {
                    RangeEditAction.InsertCells => InsertCellsAction(rangeCommands, sessionId, sheetName, rangeAddress, shiftDirection),
                    RangeEditAction.DeleteCells => DeleteCellsAction(rangeCommands, sessionId, sheetName, rangeAddress, shiftDirection),
                    RangeEditAction.InsertRows => InsertRowsAction(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeEditAction.DeleteRows => DeleteRowsAction(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeEditAction.InsertColumns => InsertColumnsAction(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeEditAction.DeleteColumns => DeleteColumnsAction(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeEditAction.Find => FindAction(rangeCommands, sessionId, sheetName, rangeAddress, searchValue, matchCase, matchEntireCell, searchInFormulas, searchInValues),
                    RangeEditAction.Replace => ReplaceAction(rangeCommands, sessionId, sheetName, rangeAddress, searchValue, replaceValue, matchCase, matchEntireCell, searchInFormulas, searchInValues, replaceAll),
                    RangeEditAction.Sort => SortAction(rangeCommands, sessionId, sheetName, rangeAddress, sortColumns, hasHeaderRow),
                    _ => throw new ArgumentException(
                        $"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string InsertCellsAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, string? shiftDirection)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "insert-cells");
        if (string.IsNullOrEmpty(shiftDirection))
            ExcelToolsBase.ThrowMissingParameter("shiftDirection", "insert-cells");

        if (!Enum.TryParse<InsertShiftDirection>(shiftDirection, true, out var direction))
            throw new ArgumentException($"Invalid shiftDirection '{shiftDirection}'. Must be 'Down' or 'Right'.", nameof(shiftDirection));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.InsertCells(batch, sheetName ?? "", rangeAddress!, direction));

        return JsonSerializer.Serialize(new { result.Success, errorMessage = result.ErrorMessage }, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteCellsAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, string? shiftDirection)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "delete-cells");
        if (string.IsNullOrEmpty(shiftDirection))
            ExcelToolsBase.ThrowMissingParameter("shiftDirection", "delete-cells");

        if (!Enum.TryParse<DeleteShiftDirection>(shiftDirection, true, out var direction))
            throw new ArgumentException($"Invalid shiftDirection '{shiftDirection}'. Must be 'Up' or 'Left'.", nameof(shiftDirection));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.DeleteCells(batch, sheetName ?? "", rangeAddress!, direction));

        return JsonSerializer.Serialize(new { result.Success, errorMessage = result.ErrorMessage }, ExcelToolsBase.JsonOptions);
    }

    private static string InsertRowsAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "insert-rows");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.InsertRows(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new { result.Success, errorMessage = result.ErrorMessage }, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteRowsAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "delete-rows");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.DeleteRows(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new { result.Success, errorMessage = result.ErrorMessage }, ExcelToolsBase.JsonOptions);
    }

    private static string InsertColumnsAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "insert-columns");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.InsertColumns(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new { result.Success, errorMessage = result.ErrorMessage }, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteColumnsAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "delete-columns");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.DeleteColumns(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new { result.Success, errorMessage = result.ErrorMessage }, ExcelToolsBase.JsonOptions);
    }

    private static string FindAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, string? searchValue, bool? matchCase, bool? matchEntireCell, bool? searchInFormulas, bool? searchInValues)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "find");
        if (string.IsNullOrEmpty(searchValue))
            ExcelToolsBase.ThrowMissingParameter("searchValue", "find");

        var options = new FindOptions
        {
            MatchCase = matchCase ?? false,
            MatchEntireCell = matchEntireCell ?? false,
            SearchFormulas = searchInFormulas ?? true,
            SearchValues = searchInValues ?? true
        };

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Find(batch, sheetName ?? "", rangeAddress!, searchValue!, options));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            sheetName = result.SheetName,
            rangeAddress = result.RangeAddress,
            searchValue = result.SearchValue,
            matchingCells = result.MatchingCells.Take(10).ToList(),
            totalMatches = result.MatchingCells.Count,
            errorMessage = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ReplaceAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, string? searchValue, string? replaceValue, bool? matchCase, bool? matchEntireCell, bool? searchInFormulas, bool? searchInValues, bool? replaceAll)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "replace");
        if (string.IsNullOrEmpty(searchValue))
            ExcelToolsBase.ThrowMissingParameter("searchValue", "replace");
        if (replaceValue == null)
            ExcelToolsBase.ThrowMissingParameter("replaceValue", "replace");

        var options = new ReplaceOptions
        {
            MatchCase = matchCase ?? false,
            MatchEntireCell = matchEntireCell ?? false,
            SearchFormulas = searchInFormulas ?? true,
            SearchValues = searchInValues ?? true,
            ReplaceAll = replaceAll ?? true
        };

        ExcelToolsBase.WithSession<object?>(
            sessionId,
            batch =>
            {
                commands.Replace(batch, sheetName ?? "", rangeAddress!, searchValue!, replaceValue!, options);
                return null;
            });

        return JsonSerializer.Serialize(new { Success = true }, ExcelToolsBase.JsonOptions);
    }

    private static string SortAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, List<SortColumn>? sortColumns, bool? hasHeaderRow)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "sort");
        if (sortColumns == null || sortColumns.Count == 0)
            ExcelToolsBase.ThrowMissingParameter("sortColumns", "sort");

        ExcelToolsBase.WithSession<object?>(
            sessionId,
            batch =>
            {
                commands.Sort(batch, sheetName ?? "", rangeAddress!, sortColumns!, hasHeaderRow ?? true);
                return null;
            });

        return JsonSerializer.Serialize(new { Success = true }, ExcelToolsBase.JsonOptions);
    }
}
