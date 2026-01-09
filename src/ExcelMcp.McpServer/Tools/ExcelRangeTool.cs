using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for core Excel range operations - values, formulas, copy, clear, discovery.
/// Use excel_range_edit for insert/delete/find/sort. Use excel_range_format for styling/validation.
/// Use excel_range_link for hyperlinks and cell protection.
/// </summary>
[McpServerToolType]
public static partial class ExcelRangeTool
{
    /// <summary>
    /// Core range ops: get/set values/formulas, copy, clear, discovery.
    /// DATA FORMAT: vals/fmls are 2D JSON arrays [[r1c1,r1c2],[r2c1,r2c2]].
    /// REQUIRED: sn (sheet name) + addr (e.g., 'A1:D10') for cell operations.
    /// </summary>
    /// <param name="action">Action</param>
    /// <param name="path">File path</param>
    /// <param name="sid">Session ID</param>
    /// <param name="sn">Sheet name (REQUIRED for cell addresses, empty only for named ranges)</param>
    /// <param name="addr">Range address (e.g., 'A1:D10') or named range</param>
    /// <param name="vals">2D values [[1,2],[3,4]]</param>
    /// <param name="fmls">2D formulas [["=A1+B1"]]</param>
    /// <param name="srcSn">Source sheet (copy ops)</param>
    /// <param name="srcAddr">Source range (copy ops)</param>
    /// <param name="tgtSn">Target sheet (copy ops)</param>
    /// <param name="tgtAddr">Target range (copy ops)</param>
    /// <param name="fmt">Format code (US locale)</param>
    /// <param name="fmts">2D format codes</param>
    /// <param name="cell">Cell address (get-current-region)</param>
    [McpServerTool(Name = "excel_range", Title = "Excel Range Operations")]
    [McpMeta("category", "data")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelRange(
        RangeAction action,
        string path,
        string sid,
        [DefaultValue(null)] string? sn,
        [DefaultValue(null)] string? addr,
        [DefaultValue(null)] List<List<object?>>? vals,
        [DefaultValue(null)] List<List<string>>? fmls,
        [DefaultValue(null)] string? srcSn,
        [DefaultValue(null)] string? srcAddr,
        [DefaultValue(null)] string? tgtSn,
        [DefaultValue(null)] string? tgtAddr,
        [DefaultValue(null)] string? fmt,
        [DefaultValue(null)] List<List<string>>? fmts,
        [DefaultValue(null)] string? cell)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_range",
            action.ToActionString(),
            path,
            () =>
            {
                var rangeCommands = new RangeCommands();

                return action switch
                {
                    RangeAction.GetValues => GetValuesAsync(rangeCommands, sid, sn, addr),
                    RangeAction.SetValues => SetValuesAsync(rangeCommands, sid, sn, addr, vals),
                    RangeAction.GetFormulas => GetFormulasAsync(rangeCommands, sid, sn, addr),
                    RangeAction.SetFormulas => SetFormulasAsync(rangeCommands, sid, sn, addr, fmls),
                    RangeAction.GetNumberFormats => GetNumberFormatsAsync(rangeCommands, sid, sn, addr),
                    RangeAction.SetNumberFormat => SetNumberFormatAsync(rangeCommands, sid, sn, addr, fmt),
                    RangeAction.SetNumberFormats => SetNumberFormatsAsync(rangeCommands, sid, sn, addr, fmts),
                    RangeAction.ClearAll => ClearAllAsync(rangeCommands, sid, sn, addr),
                    RangeAction.ClearContents => ClearContentsAsync(rangeCommands, sid, sn, addr),
                    RangeAction.ClearFormats => ClearFormatsAsync(rangeCommands, sid, sn, addr),
                    RangeAction.Copy => CopyAsync(rangeCommands, sid, srcSn, srcAddr, tgtSn, tgtAddr),
                    RangeAction.CopyValues => CopyValuesAsync(rangeCommands, sid, srcSn, srcAddr, tgtSn, tgtAddr),
                    RangeAction.CopyFormulas => CopyFormulasAsync(rangeCommands, sid, srcSn, srcAddr, tgtSn, tgtAddr),
                    RangeAction.GetUsedRange => GetUsedRangeAsync(rangeCommands, sid, sn),
                    RangeAction.GetCurrentRegion => GetCurrentRegionAsync(rangeCommands, sid, sn, cell),
                    RangeAction.GetInfo => GetRangeInfoAsync(rangeCommands, sid, sn, addr),
                    _ => throw new ArgumentException(
                        $"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    // === VALUE OPERATIONS ===

    private static string GetValuesAsync(RangeCommands commands, string sid, string? sn, string? addr)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "get-values");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.GetValues(batch, sn ?? "", addr!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            sn = result.SheetName,
            addr = result.RangeAddress,
            v = result.Values,
            rows = result.RowCount,
            cols = result.ColumnCount,
            err = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string SetValuesAsync(RangeCommands commands, string sid, string? sn, string? addr, List<List<object?>>? vals)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "set-values");
        if (vals == null || vals.Count == 0)
            ExcelToolsBase.ThrowMissingParameter("vals", "set-values");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.SetValues(batch, sn ?? "", addr!, vals!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            err = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === FORMULA OPERATIONS ===

    private static string GetFormulasAsync(RangeCommands commands, string sid, string? sn, string? addr)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "get-formulas");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.GetFormulas(batch, sn ?? "", addr!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            sn = result.SheetName,
            addr = result.RangeAddress,
            fmls = result.Formulas,
            v = result.Values,
            rows = result.RowCount,
            cols = result.ColumnCount,
            err = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string SetFormulasAsync(RangeCommands commands, string sid, string? sn, string? addr, List<List<string>>? fmls)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "set-formulas");
        if (fmls == null || fmls.Count == 0)
            ExcelToolsBase.ThrowMissingParameter("fmls", "set-formulas");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.SetFormulas(batch, sn ?? "", addr!, fmls!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            err = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === NUMBER FORMAT OPERATIONS ===

    private static string GetNumberFormatsAsync(RangeCommands commands, string sid, string? sn, string? addr)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "get-number-formats");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.GetNumberFormats(batch, sn ?? "", addr!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            sn = result.SheetName,
            addr = result.RangeAddress,
            fmts = result.Formats,
            rows = result.RowCount,
            cols = result.ColumnCount,
            err = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string SetNumberFormatAsync(RangeCommands commands, string sid, string? sn, string? addr, string? fmt)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "set-number-format");
        if (string.IsNullOrEmpty(fmt))
            ExcelToolsBase.ThrowMissingParameter("fmt", "set-number-format");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.SetNumberFormat(batch, sn ?? "", addr!, fmt!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            err = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string SetNumberFormatsAsync(RangeCommands commands, string sid, string? sn, string? addr, List<List<string>>? fmts)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "set-number-formats");
        if (fmts == null || fmts.Count == 0)
            ExcelToolsBase.ThrowMissingParameter("fmts", "set-number-formats");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.SetNumberFormats(batch, sn ?? "", addr!, fmts!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            err = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === CLEAR OPERATIONS ===

    private static string ClearAllAsync(RangeCommands commands, string sid, string? sn, string? addr)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "clear-all");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.ClearAll(batch, sn ?? "", addr!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            err = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ClearContentsAsync(RangeCommands commands, string sid, string? sn, string? addr)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "clear-contents");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.ClearContents(batch, sn ?? "", addr!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            err = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ClearFormatsAsync(RangeCommands commands, string sid, string? sn, string? addr)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "clear-formats");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.ClearFormats(batch, sn ?? "", addr!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            err = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === COPY OPERATIONS ===

    private static string CopyAsync(RangeCommands commands, string sid, string? srcSn, string? srcAddr, string? tgtSn, string? tgtAddr)
    {
        if (string.IsNullOrEmpty(srcAddr))
            ExcelToolsBase.ThrowMissingParameter("srcAddr", "copy");
        if (string.IsNullOrEmpty(tgtAddr))
            ExcelToolsBase.ThrowMissingParameter("tgtAddr", "copy");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.Copy(batch, srcSn ?? "", srcAddr!, tgtSn ?? "", tgtAddr!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            err = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string CopyValuesAsync(RangeCommands commands, string sid, string? srcSn, string? srcAddr, string? tgtSn, string? tgtAddr)
    {
        if (string.IsNullOrEmpty(srcAddr))
            ExcelToolsBase.ThrowMissingParameter("srcAddr", "copy-values");
        if (string.IsNullOrEmpty(tgtAddr))
            ExcelToolsBase.ThrowMissingParameter("tgtAddr", "copy-values");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.CopyValues(batch, srcSn ?? "", srcAddr!, tgtSn ?? "", tgtAddr!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            err = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string CopyFormulasAsync(RangeCommands commands, string sid, string? srcSn, string? srcAddr, string? tgtSn, string? tgtAddr)
    {
        if (string.IsNullOrEmpty(srcAddr))
            ExcelToolsBase.ThrowMissingParameter("srcAddr", "copy-formulas");
        if (string.IsNullOrEmpty(tgtAddr))
            ExcelToolsBase.ThrowMissingParameter("tgtAddr", "copy-formulas");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.CopyFormulas(batch, srcSn ?? "", srcAddr!, tgtSn ?? "", tgtAddr!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            err = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    // === DISCOVERY OPERATIONS ===

    private static string GetUsedRangeAsync(RangeCommands commands, string sid, string? sn)
    {
        if (string.IsNullOrEmpty(sn))
            ExcelToolsBase.ThrowMissingParameter("sn", "get-used-range");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.GetUsedRange(batch, sn!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            sn = result.SheetName,
            addr = result.RangeAddress,
            v = result.Values,
            rows = result.RowCount,
            cols = result.ColumnCount,
            err = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string GetCurrentRegionAsync(RangeCommands commands, string sid, string? sn, string? cell)
    {
        if (string.IsNullOrEmpty(sn))
            ExcelToolsBase.ThrowMissingParameter("sn", "get-current-region");
        if (string.IsNullOrEmpty(cell))
            ExcelToolsBase.ThrowMissingParameter("cell", "get-current-region");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.GetCurrentRegion(batch, sn!, cell!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            sn = result.SheetName,
            addr = result.RangeAddress,
            v = result.Values,
            rows = result.RowCount,
            cols = result.ColumnCount,
            err = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string GetRangeInfoAsync(RangeCommands commands, string sid, string? sn, string? addr)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "get-info");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.GetInfo(batch, sn ?? "", addr!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            sn = result.SheetName,
            addr = result.Address,
            rows = result.RowCount,
            cols = result.ColumnCount,
            fmt = result.NumberFormat,
            err = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }
}

