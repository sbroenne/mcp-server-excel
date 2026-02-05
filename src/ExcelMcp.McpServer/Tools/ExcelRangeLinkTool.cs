using System.ComponentModel;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel range links and protection - hyperlinks, cell locking.
/// </summary>
[McpServerToolType]
public static partial class ExcelRangeLinkTool
{
    /// <summary>
    /// Hyperlink and cell protection operations for Excel ranges.
    ///
    /// HYPERLINKS:
    /// - 'add-hyperlink': Add a clickable hyperlink to a cell (URL can be web, file, or mailto)
    /// - 'remove-hyperlink': Remove hyperlink(s) from cells while keeping the cell content
    /// - 'list-hyperlinks': Get all hyperlinks on a worksheet
    /// - 'get-hyperlink': Get hyperlink details for a specific cell
    ///
    /// CELL PROTECTION:
    /// - 'set-cell-lock': Lock or unlock cells (only effective when sheet protection is enabled)
    /// - 'get-cell-lock': Check if cells are locked
    ///
    /// Note: Cell locking only takes effect when the worksheet is protected.
    /// </summary>
    /// <param name="action">The link/protection operation to perform</param>
    /// <param name="path">Full path to Excel file (for reference/logging)</param>
    /// <param name="sessionId">Session ID from excel_file 'open'. Required for all actions.</param>
    /// <param name="sheetName">Name of the worksheet. Required for hyperlink actions. Optional for cell lock (uses active sheet if empty).</param>
    /// <param name="rangeAddress">Range address for multi-cell operations. Required for: remove-hyperlink, set-cell-lock, get-cell-lock</param>
    /// <param name="cellAddress">Single cell address. Required for: add-hyperlink, get-hyperlink</param>
    /// <param name="url">Hyperlink URL (web: 'https://...', file: 'file:///...', email: 'mailto:...'). Required for: add-hyperlink</param>
    /// <param name="displayText">Text to display in the cell (optional, defaults to URL)</param>
    /// <param name="tooltip">Tooltip text shown on hover (optional)</param>
    /// <param name="isLocked">Lock status for cell protection. true=locked (protected), false=unlocked (editable). Required for: set-cell-lock</param>
    [McpServerTool(Name = "excel_range_link", Title = "Excel Range Link Operations", Destructive = true)]
    [McpMeta("category", "data")]
    [McpMeta("requiresSession", true)]
    public static partial string RangeLink(
        RangeLinkAction action,
        string path,
        string sessionId,
        [DefaultValue(null)] string? sheetName,
        [DefaultValue(null)] string? rangeAddress,
        [DefaultValue(null)] string? cellAddress,
        [DefaultValue(null)] string? url,
        [DefaultValue(null)] string? displayText,
        [DefaultValue(null)] string? tooltip,
        [DefaultValue(null)] bool? isLocked)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_range_link",
            ServiceRegistry.RangeLink.ToActionString(action),
            path,
            () => action switch
            {
                RangeLinkAction.AddHyperlink => ForwardAddHyperlink(sessionId, sheetName, cellAddress, url, displayText, tooltip),
                RangeLinkAction.RemoveHyperlink => ForwardRemoveHyperlink(sessionId, sheetName, rangeAddress),
                RangeLinkAction.ListHyperlinks => ForwardListHyperlinks(sessionId, sheetName),
                RangeLinkAction.GetHyperlink => ForwardGetHyperlink(sessionId, sheetName, cellAddress),
                RangeLinkAction.SetCellLock => ForwardSetCellLock(sessionId, sheetName, rangeAddress, isLocked),
                RangeLinkAction.GetCellLock => ForwardGetCellLock(sessionId, sheetName, rangeAddress),
                _ => throw new ArgumentException($"Unknown action: {action} ({ServiceRegistry.RangeLink.ToActionString(action)})", nameof(action))
            });
    }

    private static string ForwardAddHyperlink(string sessionId, string? sheetName, string? cellAddress, string? url, string? displayText, string? tooltip)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "add-hyperlink");
        if (string.IsNullOrEmpty(cellAddress))
            ExcelToolsBase.ThrowMissingParameter("cellAddress", "add-hyperlink");
        if (string.IsNullOrEmpty(url))
            ExcelToolsBase.ThrowMissingParameter("url", "add-hyperlink");

        return ExcelToolsBase.ForwardToService("range.add-hyperlink", sessionId, new
        {
            sheetName,
            cellAddress,
            url,
            displayText,
            tooltip
        });
    }

    private static string ForwardRemoveHyperlink(string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "remove-hyperlink");
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "remove-hyperlink");

        return ExcelToolsBase.ForwardToService("range.remove-hyperlink", sessionId, new
        {
            sheetName,
            range = rangeAddress
        });
    }

    private static string ForwardListHyperlinks(string sessionId, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "list-hyperlinks");

        return ExcelToolsBase.ForwardToService("range.list-hyperlinks", sessionId, new
        {
            sheetName
        });
    }

    private static string ForwardGetHyperlink(string sessionId, string? sheetName, string? cellAddress)
    {
        if (string.IsNullOrEmpty(sheetName))
            ExcelToolsBase.ThrowMissingParameter("sheetName", "get-hyperlink");
        if (string.IsNullOrEmpty(cellAddress))
            ExcelToolsBase.ThrowMissingParameter("cellAddress", "get-hyperlink");

        return ExcelToolsBase.ForwardToService("range.get-hyperlink", sessionId, new
        {
            sheetName,
            cellAddress
        });
    }

    private static string ForwardSetCellLock(string sessionId, string? sheetName, string? rangeAddress, bool? isLocked)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "set-cell-lock");
        if (isLocked == null)
            ExcelToolsBase.ThrowMissingParameter("isLocked", "set-cell-lock");

        return ExcelToolsBase.ForwardToService("range.set-cell-lock", sessionId, new
        {
            sheetName = sheetName ?? "",
            range = rangeAddress,
            locked = isLocked
        });
    }

    private static string ForwardGetCellLock(string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-cell-lock");

        return ExcelToolsBase.ForwardToService("range.get-cell-lock", sessionId, new
        {
            sheetName = sheetName ?? "",
            range = rangeAddress
        });
    }
}




