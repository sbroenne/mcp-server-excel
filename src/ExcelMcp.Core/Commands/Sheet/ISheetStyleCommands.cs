using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet styling operations for tab colors and visibility.
/// Use sheet for lifecycle operations (create, rename, copy, delete, move).
///
/// TAB COLORS: Use RGB values (0-255 each) to set custom tab colors for visual organization.
///
/// VISIBILITY LEVELS:
/// - 'visible': Normal visible sheet
/// - 'hidden': Hidden but accessible via Format > Sheet > Unhide
/// - 'veryhidden': Only accessible via VBA (protection against casual unhiding)
/// </summary>
[ServiceCategory("sheet", "SheetStyle")]
[McpTool("worksheet_style", Title = "Worksheet Style Operations", Destructive = true, Category = "structure",
    Description = "Worksheet styling: tab colors and visibility. TAB COLORS: RGB values 0-255 each for custom tab colors. VISIBILITY: visible (normal), hidden (accessible via Format > Sheet > Unhide), veryhidden (only accessible via VBA). Use worksheet for lifecycle operations.")]
public interface ISheetStyleCommands
{
    // === TAB COLOR OPERATIONS ===

    /// <summary>
    /// Sets the tab color for a worksheet using RGB values (0-255 each).
    /// Excel uses BGR format internally, conversion is handled automatically.
    /// Throws exception on error.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet to color</param>
    /// <param name="red">Red color component (0-255)</param>
    /// <param name="green">Green color component (0-255)</param>
    /// <param name="blue">Blue color component (0-255)</param>
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
    /// <param name="visibility">Visibility level: 'visible', 'hidden', or 'veryhidden'</param>
    [ServiceAction("set-visibility")]
    void SetVisibility(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter]
        [FromString] SheetVisibility visibility);

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
