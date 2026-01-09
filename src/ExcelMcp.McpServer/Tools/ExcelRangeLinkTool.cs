using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel range links and protection - hyperlinks, cell locking.
/// </summary>
[McpServerToolType]
public static partial class ExcelRangeLinkTool
{
    /// <summary>
    /// Range link ops: add/remove/list/get hyperlinks, cell lock/unlock.
    /// </summary>
    /// <param name="action">Action</param>
    /// <param name="path">File path</param>
    /// <param name="sid">Session ID</param>
    /// <param name="sn">Sheet name</param>
    /// <param name="addr">Range address</param>
    /// <param name="cell">Cell address (for single-cell ops)</param>
    /// <param name="url">Hyperlink URL</param>
    /// <param name="text">Display text</param>
    /// <param name="tip">Tooltip</param>
    /// <param name="locked">Lock status (true=locked, false=unlocked)</param>
    [McpServerTool(Name = "excel_range_link", Title = "Excel Range Link Operations")]
    [McpMeta("category", "data")]
    [McpMeta("requiresSession", true)]
    public static partial string RangeLink(
        RangeLinkAction action,
        string path,
        string sid,
        [DefaultValue(null)] string? sn,
        [DefaultValue(null)] string? addr,
        [DefaultValue(null)] string? cell,
        [DefaultValue(null)] string? url,
        [DefaultValue(null)] string? text,
        [DefaultValue(null)] string? tip,
        [DefaultValue(null)] bool? locked)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_range_link",
            action.ToActionString(),
            path,
            () =>
            {
                var rangeCommands = new RangeCommands();

                return action switch
                {
                    RangeLinkAction.AddHyperlink => AddHyperlinkAsync(rangeCommands, sid, sn, cell, url, text, tip),
                    RangeLinkAction.RemoveHyperlink => RemoveHyperlinkAsync(rangeCommands, sid, sn, addr),
                    RangeLinkAction.ListHyperlinks => ListHyperlinksAsync(rangeCommands, sid, sn),
                    RangeLinkAction.GetHyperlink => GetHyperlinkAsync(rangeCommands, sid, sn, cell),
                    RangeLinkAction.SetCellLock => SetCellLockAsync(rangeCommands, sid, sn, addr, locked),
                    RangeLinkAction.GetCellLock => GetCellLockAsync(rangeCommands, sid, sn, addr),
                    _ => throw new ArgumentException(
                        $"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string AddHyperlinkAsync(RangeCommands commands, string sid, string? sn, string? cell, string? url, string? text, string? tip)
    {
        if (string.IsNullOrEmpty(sn))
            ExcelToolsBase.ThrowMissingParameter("sn", "add-hyperlink");
        if (string.IsNullOrEmpty(cell))
            ExcelToolsBase.ThrowMissingParameter("cell", "add-hyperlink");
        if (string.IsNullOrEmpty(url))
            ExcelToolsBase.ThrowMissingParameter("url", "add-hyperlink");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.AddHyperlink(batch, sn!, cell!, url!, text, tip));

        return JsonSerializer.Serialize(new { result.Success, err = result.ErrorMessage }, ExcelToolsBase.JsonOptions);
    }

    private static string RemoveHyperlinkAsync(RangeCommands commands, string sid, string? sn, string? addr)
    {
        if (string.IsNullOrEmpty(sn))
            ExcelToolsBase.ThrowMissingParameter("sn", "remove-hyperlink");
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "remove-hyperlink");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.RemoveHyperlink(batch, sn!, addr!));

        return JsonSerializer.Serialize(new { result.Success, err = result.ErrorMessage }, ExcelToolsBase.JsonOptions);
    }

    private static string ListHyperlinksAsync(RangeCommands commands, string sid, string? sn)
    {
        if (string.IsNullOrEmpty(sn))
            ExcelToolsBase.ThrowMissingParameter("sn", "list-hyperlinks");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.ListHyperlinks(batch, sn!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            sn = ((dynamic)result).SheetName,
            links = ((dynamic)result).Hyperlinks,
            err = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string GetHyperlinkAsync(RangeCommands commands, string sid, string? sn, string? cell)
    {
        if (string.IsNullOrEmpty(sn))
            ExcelToolsBase.ThrowMissingParameter("sn", "get-hyperlink");
        if (string.IsNullOrEmpty(cell))
            ExcelToolsBase.ThrowMissingParameter("cell", "get-hyperlink");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.GetHyperlink(batch, sn!, cell!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            cell = ((dynamic)result).CellAddress,
            url = ((dynamic)result).Url,
            text = ((dynamic)result).DisplayText,
            tip = ((dynamic)result).Tooltip,
            err = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string SetCellLockAsync(RangeCommands commands, string sid, string? sn, string? addr, bool? locked)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "set-cell-lock");
        if (locked == null)
            ExcelToolsBase.ThrowMissingParameter("locked", "set-cell-lock");

        ExcelToolsBase.WithSession<object?>(
            sid,
            batch =>
            {
                commands.SetCellLock(batch, sn ?? "", addr!, locked!.Value);
                return null;
            });

        return JsonSerializer.Serialize(new { Success = true }, ExcelToolsBase.JsonOptions);
    }

    private static string GetCellLockAsync(RangeCommands commands, string sid, string? sn, string? addr)
    {
        if (string.IsNullOrEmpty(addr))
            ExcelToolsBase.ThrowMissingParameter("addr", "get-cell-lock");

        var result = ExcelToolsBase.WithSession(
            sid,
            batch => commands.GetCellLock(batch, sn ?? "", addr!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            locked = ((dynamic)result).Locked,
            err = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }
}
