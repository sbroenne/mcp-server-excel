using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Hyperlink and cell protection operations for Excel ranges.
/// Use range for values/formulas, rangeformat for styling.
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
[ServiceCategory("rangelink", "RangeLink")]
[McpTool("range_link", Title = "Range Link Operations", Destructive = true, Category = "data",
    Description = "Hyperlink and cell protection operations. HYPERLINKS: add-hyperlink (URL: web, file, mailto), remove-hyperlink (keeps cell content), list-hyperlinks (all on worksheet), get-hyperlink (specific cell). CELL PROTECTION: set-cell-lock/get-cell-lock (only effective when sheet protection is enabled).")]
public interface IRangeLinkCommands
{
    // === HYPERLINK OPERATIONS ===

    /// <summary>
    /// Adds hyperlink to a single cell.
    /// Excel COM: Worksheet.Hyperlinks.Add()
    /// </summary>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="cellAddress">Single cell address (e.g., 'A1')</param>
    /// <param name="url">Hyperlink URL (web: 'https://...', file: 'file:///...', email: 'mailto:...')</param>
    /// <param name="displayText">Text to display in the cell (optional, defaults to URL)</param>
    /// <param name="tooltip">Tooltip text shown on hover (optional)</param>
    [ServiceAction("add-hyperlink")]
    OperationResult AddHyperlink(IExcelBatch batch, string sheetName, [RequiredParameter] string cellAddress, [RequiredParameter] string url, string? displayText = null, string? tooltip = null);

    /// <summary>
    /// Removes hyperlink from a single cell or all hyperlinks from a range.
    /// Excel COM: Range.Hyperlinks.Delete()
    /// </summary>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Cell range address to remove hyperlinks from (e.g., 'A1:D10')</param>
    [ServiceAction("remove-hyperlink")]
    OperationResult RemoveHyperlink(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Lists all hyperlinks in a worksheet.
    /// Excel COM: Worksheet.Hyperlinks collection
    /// </summary>
    /// <param name="sheetName">Name of the worksheet</param>
    [ServiceAction("list-hyperlinks")]
    RangeHyperlinkResult ListHyperlinks(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Gets hyperlink from a specific cell.
    /// Excel COM: Range.Hyperlink
    /// </summary>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="cellAddress">Single cell address (e.g., 'A1')</param>
    [ServiceAction("get-hyperlink")]
    RangeHyperlinkResult GetHyperlink(IExcelBatch batch, string sheetName, [RequiredParameter] string cellAddress);

    // === CELL PROTECTION OPERATIONS ===

    /// <summary>
    /// Locks or unlocks cells (requires worksheet protection to take effect).
    /// Excel COM: Range.Locked
    /// </summary>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1:D10')</param>
    /// <param name="locked">Lock status: true = locked (protected when sheet protection enabled), false = unlocked (editable)</param>
    [ServiceAction("set-cell-lock")]
    void SetCellLock(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] bool locked);

    /// <summary>
    /// Gets lock status of first cell in range.
    /// Excel COM: Range.Locked
    /// </summary>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1:D10')</param>
    [ServiceAction("get-cell-lock")]
    RangeLockInfoResult GetCellLock(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);
}
