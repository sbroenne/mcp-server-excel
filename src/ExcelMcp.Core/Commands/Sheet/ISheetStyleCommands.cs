using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet styling and appearance commands.
/// - Tab colors: set, get, clear worksheet tab colors
/// - Visibility: show, hide, very-hide worksheets
/// Lifecycle operations (create, rename, copy, delete) are in ISheetCommands.
/// </summary>
[ServiceCategory("sheet", "SheetStyle")]
[McpTool("excel_worksheet_style")]
public interface ISheetStyleCommands
{
    // === TAB COLOR OPERATIONS ===

    /// <summary>
    /// Sets the tab color for a worksheet using RGB values (0-255 each).
    /// Excel uses BGR format internally, conversion is handled automatically.
    /// Throws exception on error.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="red">Red component (0-255)</param>
    /// <param name="green">Green component (0-255)</param>
    /// <param name="blue">Blue component (0-255)</param>
    [ServiceAction("set-tab-color")]
    void SetTabColor(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] int red,
        [RequiredParameter] int green,
        [RequiredParameter] int blue);

    /// <summary>
    /// Gets the tab color for a worksheet.
    /// Returns RGB values and hex color, or HasColor=false if no color is set.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    [ServiceAction("get-tab-color")]
    TabColorResult GetTabColor(IExcelBatch batch, [RequiredParameter] string sheetName);

    /// <summary>
    /// Clears the tab color for a worksheet (resets to default).
    /// Throws exception on error.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    [ServiceAction("clear-tab-color")]
    void ClearTabColor(IExcelBatch batch, [RequiredParameter] string sheetName);

    // === VISIBILITY OPERATIONS ===

    /// <summary>
    /// Sets worksheet visibility level.
    /// - visible: Normal visible state
    /// - hidden: Hidden via UI, user can unhide
    /// - veryhidden: Requires code to unhide (security/protection)
    /// Throws exception on error.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="visibility">Visibility level: visible, hidden, or veryhidden</param>
    [ServiceAction("set-visibility")]
    void SetVisibility(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] SheetVisibility visibility);

    /// <summary>
    /// Gets worksheet visibility level
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    [ServiceAction("get-visibility")]
    SheetVisibilityResult GetVisibility(IExcelBatch batch, [RequiredParameter] string sheetName);

    /// <summary>
    /// Shows a hidden or very hidden worksheet.
    /// Convenience method equivalent to SetVisibility(..., SheetVisibility.Visible).
    /// Throws exception on error.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    [ServiceAction("show")]
    void Show(IExcelBatch batch, [RequiredParameter] string sheetName);

    /// <summary>
    /// Hides a worksheet (user can unhide via Excel UI).
    /// Convenience method equivalent to SetVisibility(..., SheetVisibility.Hidden).
    /// Throws exception on error.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    [ServiceAction("hide")]
    void Hide(IExcelBatch batch, [RequiredParameter] string sheetName);

    /// <summary>
    /// Very hides a worksheet (requires code to unhide, for protection).
    /// Convenience method equivalent to SetVisibility(..., SheetVisibility.VeryHidden).
    /// Throws exception on error.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    [ServiceAction("very-hide")]
    void VeryHide(IExcelBatch batch, [RequiredParameter] string sheetName);
}
