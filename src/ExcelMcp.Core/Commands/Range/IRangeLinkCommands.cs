using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Manage hyperlinks (add, remove, list) and cell lock state for worksheet protection.
/// Use range for values/formulas, rangeformat for styling.
/// </summary>
[ServiceCategory("rangelink", "RangeLink")]
[McpTool("excel_range_link")]
public interface IRangeLinkCommands
{
    // === HYPERLINK OPERATIONS ===

    /// <summary>
    /// Adds hyperlink to a single cell.
    /// Excel COM: Worksheet.Hyperlinks.Add()
    /// </summary>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="cellAddress">Cell address (e.g., A1)</param>
    /// <param name="url">URL or file path for the hyperlink</param>
    /// <param name="displayText">Optional text to display (defaults to URL)</param>
    /// <param name="tooltip">Optional tooltip on hover</param>
    [ServiceAction("add-hyperlink")]
    OperationResult AddHyperlink(IExcelBatch batch, string sheetName, [RequiredParameter] string cellAddress, [RequiredParameter] string url, string? displayText = null, string? tooltip = null);

    /// <summary>
    /// Removes hyperlink from a single cell or all hyperlinks from a range.
    /// Excel COM: Range.Hyperlinks.Delete()
    /// </summary>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range address to remove hyperlinks from</param>
    [ServiceAction("remove-hyperlink")]
    OperationResult RemoveHyperlink(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Lists all hyperlinks in a worksheet.
    /// Excel COM: Worksheet.Hyperlinks collection
    /// </summary>
    /// <param name="sheetName">Worksheet name</param>
    [ServiceAction("list-hyperlinks")]
    RangeHyperlinkResult ListHyperlinks(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Gets hyperlink from a specific cell.
    /// Excel COM: Range.Hyperlink
    /// </summary>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="cellAddress">Cell address (e.g., A1)</param>
    [ServiceAction("get-hyperlink")]
    RangeHyperlinkResult GetHyperlink(IExcelBatch batch, string sheetName, [RequiredParameter] string cellAddress);

    // === CELL PROTECTION OPERATIONS ===

    /// <summary>
    /// Locks or unlocks cells (requires worksheet protection to take effect).
    /// Excel COM: Range.Locked
    /// </summary>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range address</param>
    /// <param name="locked">True to lock cells, false to unlock</param>
    [ServiceAction("set-cell-lock")]
    void SetCellLock(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] bool locked);

    /// <summary>
    /// Gets lock status of first cell in range.
    /// Excel COM: Range.Locked
    /// </summary>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range address</param>
    [ServiceAction("get-cell-lock")]
    RangeLockInfoResult GetCellLock(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);
}
