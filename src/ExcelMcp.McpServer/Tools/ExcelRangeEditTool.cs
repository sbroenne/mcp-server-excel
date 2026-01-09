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
    /// Range edit ops: insert/delete cells/rows/cols, find/replace, sort.
    /// </summary>
    /// <param name="action">Action</param>
    /// <param name="path">File path</param>
    /// <param name="sid">Session ID</param>
    /// <param name="sn">Sheet name</param>
    /// <param name="addr">Range address</param>
    /// <param name="shift">Shift direction: Down, Right, Up, Left</param>
    /// <param name="search">Search value</param>
    /// <param name="repl">Replace value</param>
    /// <param name="matchCase">Match case (default: false)</param>
    /// <param name="matchAll">Match entire cell (default: false)</param>
    /// <param name="inFmls">Search in formulas (default: true)</param>
    /// <param name="inVals">Search in values (default: true)</param>
    /// <param name="replAll">Replace all (default: true)</param>
    /// <param name="sortCols">Sort columns JSON [{col:1,asc:true}]</param>
    /// <param name="hasHdr">Has header row (default: true)</param>
    [McpServerTool(Name = "excel_range_edit", Title = "Excel Range Edit Operations")]
    [McpMeta("category", "data")]
    [McpMeta("requiresSession", true)]
    public static partial string RangeEdit(
        RangeEditAction action,
        string path,
        string sid,
        [DefaultValue(null)] string? sn,
        [DefaultValue(null)] string? addr,
        [DefaultValue(null)] string? shift,
        [DefaultValue(null)] string? search,
        [DefaultValue(null)] string? repl,
        [DefaultValue(null)] bool? matchCase,
        [DefaultValue(null)] bool? matchAll,
        [DefaultValue(null)] bool? inFmls,
        [DefaultValue(null)] bool? inVals,
        [DefaultValue(null)] bool? replAll,
        [DefaultValue(null)] List<SortColumn>? sortCols,
        [DefaultValue(null)] bool? hasHdr)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_range_edit",
            action.ToActionString(),
            path,
            () =>
            {
                var rangeCommands = new RangeCommands();

                return action switch
                {
                    RangeEditAction.InsertCells => InsertCellsAsync(rangeCommands, sid, sn, addr, shift),
                    RangeEditAction.DeleteCells => DeleteCellsAsync(rangeCommands, sid, sn, addr, shift),
                    RangeEditAction.InsertRows => InsertRowsAsync(rangeCommands, sid, sn, addr),
                    RangeEditAction.DeleteRows => DeleteRowsAsync(rangeCommands, sid, sn, addr),
                    RangeEditAction.InsertColumns => InsertColumnsAsync(rangeCommands, sid, sn, addr),
                    RangeEditAction.DeleteColumns => DeleteColumnsAsync(rangeCommands, sid, sn, addr),
                    RangeEditAction.Find => FindAsync(rangeCommands, sid, sn, addr, search, matchCase, matchAll, inFmls, inVals),
                    RangeEditAction.Replace => ReplaceAsync(rangeCommands, sid, sn, addr, search, repl, matchCase, matchAll, inFmls, inVals, replAll),
                    RangeEditAction.Sort => SortAsync(rangeCommands, sid, sn, addr, sortCols, hasHdr),
                    _ => throw new ArgumentException(
                        $"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string InsertCellsAsync(RangeCommands commands, string sid, string? sn, string? addr, string? shift)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "insert-cells");
        if (string.IsNullOrEmpty(shift))
            ExcelToolsBase.ThrowMissingParameter("shift", "insert-cells");

        if (!Enum.TryParse<InsertShiftDirection>(shift, true, out var shiftDirection))
            throw new ArgumentException($"Invalid shift '{shift}'. Must be 'Down' or 'Right'.", nameof(shift));

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.InsertCells(batch, sn ?? "", addr!, shiftDirection));

        return JsonSerializer.Serialize(new { result.Success, err = result.ErrorMessage }, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteCellsAsync(RangeCommands commands, string sid, string? sn, string? addr, string? shift)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "delete-cells");
        if (string.IsNullOrEmpty(shift))
            ExcelToolsBase.ThrowMissingParameter("shift", "delete-cells");

        if (!Enum.TryParse<DeleteShiftDirection>(shift, true, out var shiftDirection))
            throw new ArgumentException($"Invalid shift '{shift}'. Must be 'Up' or 'Left'.", nameof(shift));

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.DeleteCells(batch, sn ?? "", addr!, shiftDirection));

        return JsonSerializer.Serialize(new { result.Success, err = result.ErrorMessage }, ExcelToolsBase.JsonOptions);
    }

    private static string InsertRowsAsync(RangeCommands commands, string sid, string? sn, string? addr)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "insert-rows");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.InsertRows(batch, sn ?? "", addr!));

        return JsonSerializer.Serialize(new { result.Success, err = result.ErrorMessage }, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteRowsAsync(RangeCommands commands, string sid, string? sn, string? addr)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "delete-rows");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.DeleteRows(batch, sn ?? "", addr!));

        return JsonSerializer.Serialize(new { result.Success, err = result.ErrorMessage }, ExcelToolsBase.JsonOptions);
    }

    private static string InsertColumnsAsync(RangeCommands commands, string sid, string? sn, string? addr)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "insert-columns");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.InsertColumns(batch, sn ?? "", addr!));

        return JsonSerializer.Serialize(new { result.Success, err = result.ErrorMessage }, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteColumnsAsync(RangeCommands commands, string sid, string? sn, string? addr)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "delete-columns");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.DeleteColumns(batch, sn ?? "", addr!));

        return JsonSerializer.Serialize(new { result.Success, err = result.ErrorMessage }, ExcelToolsBase.JsonOptions);
    }

    private static string FindAsync(RangeCommands commands, string sid, string? sn, string? addr, string? search, bool? matchCase, bool? matchAll, bool? inFmls, bool? inVals)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "find");
        if (string.IsNullOrEmpty(search))
            ExcelToolsBase.ThrowMissingParameter("search", "find");

        var options = new FindOptions
        {
            MatchCase = matchCase ?? false,
            MatchEntireCell = matchAll ?? false,
            SearchFormulas = inFmls ?? true,
            SearchValues = inVals ?? true
        };

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.Find(batch, sn ?? "", addr!, search!, options));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            sn = result.SheetName,
            addr = result.RangeAddress,
            search = result.SearchValue,
            matches = result.MatchingCells.Take(10).ToList(),
            total = result.MatchingCells.Count,
            err = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ReplaceAsync(RangeCommands commands, string sid, string? sn, string? addr, string? search, string? repl, bool? matchCase, bool? matchAll, bool? inFmls, bool? inVals, bool? replAll)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "replace");
        if (string.IsNullOrEmpty(search))
            ExcelToolsBase.ThrowMissingParameter("search", "replace");
        if (repl == null)
            ExcelToolsBase.ThrowMissingParameter("repl", "replace");

        var options = new ReplaceOptions
        {
            MatchCase = matchCase ?? false,
            MatchEntireCell = matchAll ?? false,
            SearchFormulas = inFmls ?? true,
            SearchValues = inVals ?? true,
            ReplaceAll = replAll ?? true
        };

        ExcelToolsBase.WithSession<object?>(
            sid,
            batch =>
            {
                commands.Replace(batch, sn ?? "", addr!, search!, repl!, options);
                return null;
            });

        return JsonSerializer.Serialize(new { Success = true }, ExcelToolsBase.JsonOptions);
    }

    private static string SortAsync(RangeCommands commands, string sid, string? sn, string? addr, List<SortColumn>? sortCols, bool? hasHdr)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "sort");
        if (sortCols == null || sortCols.Count == 0)
            ExcelToolsBase.ThrowMissingParameter("sortCols", "sort");

        ExcelToolsBase.WithSession<object?>(
            sid,
            batch =>
            {
                commands.Sort(batch, sn ?? "", addr!, sortCols!, hasHdr ?? true);
                return null;
            });

        return JsonSerializer.Serialize(new { Success = true }, ExcelToolsBase.JsonOptions);
    }
}
