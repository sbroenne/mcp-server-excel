using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Hyperlink, threaded comment, and cell protection operations for Excel ranges.
/// Use range for values/formulas, rangeformat for styling.
///
/// HYPERLINKS:
/// - 'add-hyperlink': Add an external or internal workbook hyperlink
/// - 'update-hyperlink': Update the target or display metadata of an existing hyperlink
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
    Description = "Hyperlink, threaded comment, and cell protection operations. THREADED COMMENTS: add-threaded-comment, list-threaded-comments, add-threaded-comment-reply, delete-threaded-comment use the local Excel PIA. Cloud mentions, assignments, reactions, presence, and coauthoring state are not exposed by local Excel COM. HYPERLINKS: add-hyperlink creates external links with url or internal workbook links with subAddress; update-hyperlink changes an existing target, display text, or tooltip; remove-hyperlink keeps cell content; list-hyperlinks returns all worksheet links; get-hyperlink reads a specific cell. At least url or subAddress is required when adding. CELL PROTECTION: set-cell-lock/get-cell-lock only take effect when sheet protection is enabled.")]
public interface IRangeLinkCommands
{
    // === HYPERLINK OPERATIONS ===

    /// <summary>
    /// Adds hyperlink to a single cell.
    /// Excel COM: Worksheet.Hyperlinks.Add()
    /// </summary>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="cellAddress">Single cell address (e.g., 'A1')</param>
    /// <param name="url">Optional external URL or file path. Omit for an internal workbook link.</param>
    /// <param name="displayText">Text to display in the cell (optional, defaults to URL)</param>
    /// <param name="tooltip">Tooltip text shown on hover (optional)</param>
    /// <param name="subAddress">Optional internal workbook target such as "'Sheet2'!A1"</param>
    [ServiceAction("add-hyperlink")]
    OperationResult AddHyperlink(
        IExcelBatch batch,
        string sheetName,
        [RequiredParameter] string cellAddress,
        string? url = null,
        string? displayText = null,
        string? tooltip = null,
        string? subAddress = null);

    /// <summary>
    /// Updates an existing hyperlink target or display metadata in a single cell.
    /// Omitted values remain unchanged; pass an empty string to clear url, subAddress, or tooltip.
    /// </summary>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="cellAddress">Single cell address containing a hyperlink</param>
    /// <param name="url">Optional new external URL or file path</param>
    /// <param name="displayText">Optional new display text</param>
    /// <param name="tooltip">Optional new tooltip; empty string clears it</param>
    /// <param name="subAddress">Optional new workbook sub-address</param>
    [ServiceAction("update-hyperlink")]
    OperationResult UpdateHyperlink(
        IExcelBatch batch,
        string sheetName,
        [RequiredParameter] string cellAddress,
        string? url = null,
        string? displayText = null,
        string? tooltip = null,
        string? subAddress = null);

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

    // === THREADED COMMENT OPERATIONS ===

    /// <summary>
    /// Adds a top-level threaded comment to one cell.
    /// Excel COM: Range.AddCommentThreaded().
    /// </summary>
    [ServiceAction("add-threaded-comment")]
    OperationResult AddThreadedComment(
        IExcelBatch batch,
        string sheetName,
        [RequiredParameter] string cellAddress,
        [RequiredParameter] string text);

    /// <summary>
    /// Lists a cell's top-level threaded comment and its replies.
    /// Excel COM: Range.CommentThreaded and CommentThreaded.Replies.
    /// </summary>
    [ServiceAction("list-threaded-comments")]
    ThreadedCommentsResult ListThreadedComments(
        IExcelBatch batch,
        string sheetName,
        [RequiredParameter] string cellAddress);

    /// <summary>
    /// Adds a reply to a cell's existing threaded comment.
    /// Excel COM: CommentThreaded.AddReply().
    /// </summary>
    [ServiceAction("add-threaded-comment-reply")]
    OperationResult AddThreadedCommentReply(
        IExcelBatch batch,
        string sheetName,
        [RequiredParameter] string cellAddress,
        [RequiredParameter] string text);

    /// <summary>
    /// Deletes a cell's threaded comment and all replies.
    /// Excel COM: CommentThreaded.Delete().
    /// </summary>
    [ServiceAction("delete-threaded-comment")]
    OperationResult DeleteThreadedComment(
        IExcelBatch batch,
        string sheetName,
        [RequiredParameter] string cellAddress);

    // === CELL PROTECTION OPERATIONS ===

    /// <summary>
    /// Locks or unlocks cells (requires worksheet protection to take effect).
    /// Excel COM: Range.Locked
    /// </summary>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1:D10')</param>
    /// <param name="locked">Lock status: true = locked (protected when sheet protection enabled), false = unlocked (editable)</param>
    [ServiceAction("set-cell-lock")]
    OperationResult SetCellLock(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] bool locked);

    /// <summary>
    /// Gets lock status of first cell in range.
    /// Excel COM: Range.Locked
    /// </summary>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1:D10')</param>
    [ServiceAction("get-cell-lock")]
    RangeLockInfoResult GetCellLock(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);
}
