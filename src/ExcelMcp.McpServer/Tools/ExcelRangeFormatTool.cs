using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel range formatting - styles, validation, merge, autofit.
/// </summary>
[McpServerToolType]
public static partial class ExcelRangeFormatTool
{
    /// <summary>
    /// Range format ops: styles, validation, merge/unmerge, autofit.
    /// </summary>
    /// <param name="action">Action</param>
    /// <param name="path">File path</param>
    /// <param name="sid">Session ID</param>
    /// <param name="sn">Sheet name</param>
    /// <param name="addr">Range address</param>
    /// <param name="style">Style name (Heading 1, Good, Currency, etc.)</param>
    /// <param name="font">Font name</param>
    /// <param name="size">Font size</param>
    /// <param name="bold">Bold</param>
    /// <param name="italic">Italic</param>
    /// <param name="uline">Underline</param>
    /// <param name="fgClr">Font color #RRGGBB</param>
    /// <param name="bgClr">Fill color #RRGGBB</param>
    /// <param name="bdrStyle">Border style</param>
    /// <param name="bdrClr">Border color</param>
    /// <param name="bdrWt">Border weight</param>
    /// <param name="hAlign">Horizontal align</param>
    /// <param name="vAlign">Vertical align</param>
    /// <param name="wrap">Wrap text</param>
    /// <param name="orient">Text orientation degrees</param>
    /// <param name="valType">Validation type: list, whole, decimal, date, time, textLength, custom</param>
    /// <param name="valOp">Validation operator: between, equal, greaterThan, etc.</param>
    /// <param name="valFml1">Validation formula1 (for list: =$A$1:$A$10)</param>
    /// <param name="valFml2">Validation formula2 (for between/notBetween)</param>
    /// <param name="showIn">Show input message</param>
    /// <param name="inTitle">Input title</param>
    /// <param name="inMsg">Input message</param>
    /// <param name="showErr">Show error alert</param>
    /// <param name="errStyle">Error style: stop, warning, information</param>
    /// <param name="errTitle">Error title</param>
    /// <param name="errMsg">Error message</param>
    /// <param name="ignBlank">Ignore blank cells</param>
    /// <param name="showDrop">Show dropdown for list validation</param>
    [McpServerTool(Name = "excel_range_format", Title = "Excel Range Format Operations")]
    [McpMeta("category", "data")]
    [McpMeta("requiresSession", true)]
    public static partial string RangeFormat(
        RangeFormatAction action,
        string path,
        string sid,
        [DefaultValue(null)] string? sn,
        [DefaultValue(null)] string? addr,
        [DefaultValue(null)] string? style,
        [DefaultValue(null)] string? font,
        [DefaultValue(null)] double? size,
        [DefaultValue(null)] bool? bold,
        [DefaultValue(null)] bool? italic,
        [DefaultValue(null)] bool? uline,
        [DefaultValue(null)] string? fgClr,
        [DefaultValue(null)] string? bgClr,
        [DefaultValue(null)] string? bdrStyle,
        [DefaultValue(null)] string? bdrClr,
        [DefaultValue(null)] string? bdrWt,
        [DefaultValue(null)] string? hAlign,
        [DefaultValue(null)] string? vAlign,
        [DefaultValue(null)] bool? wrap,
        [DefaultValue(null)] int? orient,
        [DefaultValue(null)] string? valType,
        [DefaultValue(null)] string? valOp,
        [DefaultValue(null)] string? valFml1,
        [DefaultValue(null)] string? valFml2,
        [DefaultValue(null)] bool? showIn,
        [DefaultValue(null)] string? inTitle,
        [DefaultValue(null)] string? inMsg,
        [DefaultValue(null)] bool? showErr,
        [DefaultValue(null)] string? errStyle,
        [DefaultValue(null)] string? errTitle,
        [DefaultValue(null)] string? errMsg,
        [DefaultValue(null)] bool? ignBlank,
        [DefaultValue(null)] bool? showDrop)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_range_format",
            action.ToActionString(),
            path,
            () =>
            {
                var rangeCommands = new RangeCommands();

                return action switch
                {
                    RangeFormatAction.GetStyle => GetStyleAsync(rangeCommands, sid, sn, addr),
                    RangeFormatAction.SetStyle => SetStyleAsync(rangeCommands, sid, sn, addr, style),
                    RangeFormatAction.FormatRange => FormatRangeAsync(rangeCommands, sid, sn, addr, font, size, bold, italic, uline, fgClr, bgClr, bdrStyle, bdrClr, bdrWt, hAlign, vAlign, wrap, orient),
                    RangeFormatAction.ValidateRange => ValidateRangeAsync(rangeCommands, sid, sn, addr, valType, valOp, valFml1, valFml2, showIn, inTitle, inMsg, showErr, errStyle, errTitle, errMsg, ignBlank, showDrop),
                    RangeFormatAction.GetValidation => GetValidationAsync(rangeCommands, sid, sn, addr),
                    RangeFormatAction.RemoveValidation => RemoveValidationAsync(rangeCommands, sid, sn, addr),
                    RangeFormatAction.AutoFitColumns => AutoFitColumnsAsync(rangeCommands, sid, sn, addr),
                    RangeFormatAction.AutoFitRows => AutoFitRowsAsync(rangeCommands, sid, sn, addr),
                    RangeFormatAction.MergeCells => MergeCellsAsync(rangeCommands, sid, sn, addr),
                    RangeFormatAction.UnmergeCells => UnmergeCellsAsync(rangeCommands, sid, sn, addr),
                    RangeFormatAction.GetMergeInfo => GetMergeInfoAsync(rangeCommands, sid, sn, addr),
                    _ => throw new ArgumentException(
                        $"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string GetStyleAsync(RangeCommands commands, string sid, string? sn, string? addr)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "get-style");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.GetStyle(batch, sn ?? "", addr!));

        return JsonSerializer.Serialize(new
        {
            Success = true,
            sn,
            addr,
            style = result.StyleName,
            builtin = result.IsBuiltInStyle,
            desc = result.StyleDescription
        }, ExcelToolsBase.JsonOptions);
    }

    private static string SetStyleAsync(RangeCommands commands, string sid, string? sn, string? addr, string? style)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "set-style");
        if (string.IsNullOrEmpty(style))
            ExcelToolsBase.ThrowMissingParameter("style", "set-style");

        ExcelToolsBase.WithSession<object?>(
            sid,
            batch =>
            {
                commands.SetStyle(batch, sn ?? "", addr!, style!);
                return null;
            });

        return JsonSerializer.Serialize(new { Success = true }, ExcelToolsBase.JsonOptions);
    }

    private static string FormatRangeAsync(
        RangeCommands commands, string sid, string? sn, string? addr,
        string? font, double? size, bool? bold, bool? italic, bool? uline,
        string? fgClr, string? bgClr, string? bdrStyle, string? bdrClr, string? bdrWt,
        string? hAlign, string? vAlign, bool? wrap, int? orient)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "format-range");

        ExcelToolsBase.WithSession<object?>(
            sid,
            batch =>
            {
                commands.FormatRange(batch, sn ?? "", addr!, font, size, bold, italic, uline, fgClr, bgClr, bdrStyle, bdrClr, bdrWt, hAlign, vAlign, wrap, orient);
                return null;
            });

        return JsonSerializer.Serialize(new { Success = true }, ExcelToolsBase.JsonOptions);
    }

    private static string ValidateRangeAsync(
        RangeCommands commands, string sid, string? sn, string? addr,
        string? valType, string? valOp, string? valFml1, string? valFml2,
        bool? showIn, string? inTitle, string? inMsg,
        bool? showErr, string? errStyle, string? errTitle, string? errMsg,
        bool? ignBlank, bool? showDrop)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "validate-range");
        if (string.IsNullOrEmpty(valType))
            ExcelToolsBase.ThrowMissingParameter("valType", "validate-range");

        ExcelToolsBase.WithSession<object?>(
            sid,
            batch =>
            {
                commands.ValidateRange(batch, sn ?? "", addr!, valType!, valOp, valFml1, valFml2, showIn, inTitle, inMsg, showErr, errStyle, errTitle, errMsg, ignBlank, showDrop);
                return null;
            });

        return JsonSerializer.Serialize(new { Success = true }, ExcelToolsBase.JsonOptions);
    }

    private static string GetValidationAsync(RangeCommands commands, string sid, string? sn, string? addr)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "get-validation");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.GetValidation(batch, sn ?? "", addr!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            type = ((dynamic)result).ValidationType,
            op = ((dynamic)result).ValidationOperator,
            fml1 = ((dynamic)result).Formula1,
            fml2 = ((dynamic)result).Formula2,
            showIn = ((dynamic)result).ShowInputMessage,
            inTitle = ((dynamic)result).InputTitle,
            inMsg = ((dynamic)result).InputMessage,
            showErr = ((dynamic)result).ShowErrorAlert,
            errStyle = ((dynamic)result).ErrorStyle,
            errTitle = ((dynamic)result).ErrorTitle,
            errMsg = ((dynamic)result).ErrorMessage,
            err = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string RemoveValidationAsync(RangeCommands commands, string sid, string? sn, string? addr)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "remove-validation");

        ExcelToolsBase.WithSession<object?>(
            sid,
            batch =>
            {
                commands.RemoveValidation(batch, sn ?? "", addr!);
                return null;
            });

        return JsonSerializer.Serialize(new { Success = true }, ExcelToolsBase.JsonOptions);
    }

    private static string AutoFitColumnsAsync(RangeCommands commands, string sid, string? sn, string? addr)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "auto-fit-columns");

        ExcelToolsBase.WithSession<object?>(
            sid,
            batch =>
            {
                commands.AutoFitColumns(batch, sn ?? "", addr!);
                return null;
            });

        return JsonSerializer.Serialize(new { Success = true }, ExcelToolsBase.JsonOptions);
    }

    private static string AutoFitRowsAsync(RangeCommands commands, string sid, string? sn, string? addr)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "auto-fit-rows");

        ExcelToolsBase.WithSession<object?>(
            sid,
            batch =>
            {
                commands.AutoFitRows(batch, sn ?? "", addr!);
                return null;
            });

        return JsonSerializer.Serialize(new { Success = true }, ExcelToolsBase.JsonOptions);
    }

    private static string MergeCellsAsync(RangeCommands commands, string sid, string? sn, string? addr)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "merge-cells");

        ExcelToolsBase.WithSession<object?>(
            sid,
            batch =>
            {
                commands.MergeCells(batch, sn ?? "", addr!);
                return null;
            });

        return JsonSerializer.Serialize(new { Success = true }, ExcelToolsBase.JsonOptions);
    }

    private static string UnmergeCellsAsync(RangeCommands commands, string sid, string? sn, string? addr)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "unmerge-cells");

        ExcelToolsBase.WithSession<object?>(
            sid,
            batch =>
            {
                commands.UnmergeCells(batch, sn ?? "", addr!);
                return null;
            });

        return JsonSerializer.Serialize(new { Success = true }, ExcelToolsBase.JsonOptions);
    }

    private static string GetMergeInfoAsync(RangeCommands commands, string sid, string? sn, string? addr)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "get-merge-info");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.GetMergeInfo(batch, sn ?? "", addr!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            merged = ((dynamic)result).IsMerged,
            mergeAddr = ((dynamic)result).MergeAddress,
            err = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }
}
