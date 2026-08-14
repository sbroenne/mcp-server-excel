using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet styling operations for tab colors, visibility, and protection.
/// Use sheet for lifecycle operations (create, rename, copy, delete, move).
///
/// TAB COLORS: Use RGB values (0-255 each) to set custom tab colors for visual organization.
///
/// VISIBILITY LEVELS:
/// - 'visible': Normal visible sheet
/// - 'hidden': Hidden but accessible via Format > Sheet > Unhide
/// - 'veryhidden': Only accessible via VBA (protection against casual unhiding)
///
/// PROTECTION: Protect a worksheet to lock its contents and structure, or unprotect it.
/// </summary>
[ServiceCategory("sheet", "SheetStyle")]
[McpTool("worksheet_style", Title = "Worksheet Style Operations", Destructive = true, Category = "structure",
    Description = "Worksheet styling: tab colors, visibility, and protection. TAB COLORS: RGB values 0-255 each for custom tab colors. VISIBILITY: visible (normal), hidden (accessible via Format > Sheet > Unhide), veryhidden (only accessible via VBA). PROTECTION: protect or unprotect a worksheet. Use worksheet for lifecycle operations.")]
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
    OperationResult SetTabColor(
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
    OperationResult ClearTabColor(IExcelBatch batch, [RequiredParameter] string sheetName);

    // === PROTECTION OPERATIONS ===

    /// <summary>
    /// Protects or unprotects a worksheet.
    /// When protecting, Excel locks the sheet contents and structure unless a password is supplied.
    /// Throws exception on error.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="isProtected">Whether the worksheet should be protected</param>
    /// <param name="password">Optional password for protecting/unprotecting the sheet</param>
    [ServiceAction("set-protection")]
    OperationResult SetProtection(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] bool isProtected,
        string? password = null);

    /// <summary>
    /// Gets whether a worksheet is protected.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    [ServiceAction("get-protection")]
    SheetProtectionResult GetProtection(IExcelBatch batch, [RequiredParameter] string sheetName);

    // === CELL NOTE OPERATIONS ===

    /// <summary>
    /// Sets a legacy cell note through Excel's Comment COM API.
    /// Creates the note if one does not already exist. This is distinct from a threaded comment.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="cellAddress">Cell address such as A1</param>
    /// <param name="text">Cell note text to set</param>
    [ServiceAction("set-comment")]
    OperationResult SetComment(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string cellAddress,
        [RequiredParameter] string text);

    /// <summary>
    /// Gets legacy cell note text through Excel's Comment COM API.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="cellAddress">Cell address such as A1</param>
    [ServiceAction("get-comment")]
    SheetCommentResult GetComment(IExcelBatch batch, [RequiredParameter] string sheetName, [RequiredParameter] string cellAddress);

    /// <summary>
    /// Clears a legacy cell note through Excel's Comment COM API.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="cellAddress">Cell address such as A1</param>
    [ServiceAction("clear-comment")]
    OperationResult ClearComment(IExcelBatch batch, [RequiredParameter] string sheetName, [RequiredParameter] string cellAddress);

    // === IMAGE OPERATIONS ===

    /// <summary>
    /// Inserts an image from disk into a worksheet and anchors it to a cell.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="imagePath">Absolute path to the image file on disk</param>
    /// <param name="cellAddress">Cell address such as A1</param>
    [ServiceAction("add-image")]
    OperationResult AddImage(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string imagePath,
        [RequiredParameter] string cellAddress);

    /// <summary>
    /// Gets the number of images currently present on a worksheet.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    [ServiceAction("get-image-count")]
    WorksheetImageCountResult GetImageCount(IExcelBatch batch, [RequiredParameter] string sheetName);

    // === SHAPE OPERATIONS ===

    /// <summary>
    /// Inserts a basic rectangle shape into a worksheet and anchors it to a cell.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="cellAddress">Cell address such as A1</param>
    [ServiceAction("add-shape")]
    OperationResult AddShape(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string cellAddress);

    /// <summary>
    /// Gets the number of shapes currently present on a worksheet.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    [ServiceAction("get-shape-count")]
    WorksheetShapeCountResult GetShapeCount(IExcelBatch batch, [RequiredParameter] string sheetName);

    // === PAGE SETUP OPERATIONS ===

    /// <summary>
    /// Sets worksheet page setup properties such as orientation and fit-to-page settings.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="orientation">Page orientation: 'portrait' or 'landscape'</param>
    /// <param name="fitToPagesWide">Number of pages wide to fit the printout to</param>
    /// <param name="fitToPagesTall">Number of pages tall to fit the printout to</param>
    /// <param name="centerHorizontally">Whether to center the printout horizontally on the page</param>
    /// <param name="centerVertically">Whether to center the printout vertically on the page</param>
    [ServiceAction("set-page-setup")]
    OperationResult SetPageSetup(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string orientation,
        int? fitToPagesWide = null,
        int? fitToPagesTall = null,
        bool? centerHorizontally = null,
        bool? centerVertically = null);

    /// <summary>
    /// Reads worksheet page setup properties such as orientation and fit-to-page settings.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    [ServiceAction("get-page-setup")]
    SheetPageSetupResult GetPageSetup(IExcelBatch batch, [RequiredParameter] string sheetName);

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
    OperationResult SetVisibility(
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
    OperationResult Show(IExcelBatch batch, [RequiredParameter] string sheetName);

    /// <summary>
    /// Hides a worksheet (user can unhide via Excel UI).
    /// Convenience method equivalent to SetVisibility(..., SheetVisibility.Hidden).
    /// Throws exception on error.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    [ServiceAction("hide")]
    OperationResult Hide(IExcelBatch batch, [RequiredParameter] string sheetName);

    /// <summary>
    /// Very hides a worksheet (requires code to unhide, for protection).
    /// Convenience method equivalent to SetVisibility(..., SheetVisibility.VeryHidden).
    /// Throws exception on error.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    [ServiceAction("very-hide")]
    OperationResult VeryHide(IExcelBatch batch, [RequiredParameter] string sheetName);
}
