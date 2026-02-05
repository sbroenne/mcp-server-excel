using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Excel range hyperlink and cell protection operations.
/// Use IRangeCommands for values/formulas/copy/clear operations.
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
    [ServiceAction("add-hyperlink")]
    OperationResult AddHyperlink(IExcelBatch batch, string sheetName, [RequiredParameter] string cellAddress, [RequiredParameter] string url, string? displayText = null, string? tooltip = null);

    /// <summary>
    /// Removes hyperlink from a single cell or all hyperlinks from a range.
    /// Excel COM: Range.Hyperlinks.Delete()
    /// </summary>
    [ServiceAction("remove-hyperlink")]
    OperationResult RemoveHyperlink(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Lists all hyperlinks in a worksheet.
    /// Excel COM: Worksheet.Hyperlinks collection
    /// </summary>
    [ServiceAction("list-hyperlinks")]
    RangeHyperlinkResult ListHyperlinks(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Gets hyperlink from a specific cell.
    /// Excel COM: Range.Hyperlink
    /// </summary>
    [ServiceAction("get-hyperlink")]
    RangeHyperlinkResult GetHyperlink(IExcelBatch batch, string sheetName, [RequiredParameter] string cellAddress);

    // === CELL PROTECTION OPERATIONS ===

    /// <summary>
    /// Locks or unlocks cells (requires worksheet protection to take effect).
    /// Excel COM: Range.Locked
    /// </summary>
    [ServiceAction("set-cell-lock")]
    void SetCellLock(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] bool locked);

    /// <summary>
    /// Gets lock status of first cell in range.
    /// Excel COM: Range.Locked
    /// </summary>
    [ServiceAction("get-cell-lock")]
    RangeLockInfoResult GetCellLock(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);
}
