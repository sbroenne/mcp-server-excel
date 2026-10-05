using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet styling, visibility, protection, grouping, and outline operations.
/// Use sheet for lifecycle operations (create, rename, copy, delete, move).
///
/// TAB COLORS: Use RGB values (0-255 each) to set custom tab colors for visual organization.
///
/// VISIBILITY LEVELS:
/// - 'visible': Normal visible sheet
/// - 'hidden': Hidden but accessible via Format > Sheet > Unhide
/// - 'veryhidden': Only accessible via VBA (protection against casual unhiding)
///
/// PROTECTION: Protect selected worksheet components with explicit native permissions, or unprotect.
/// UserInterfaceOnly is runtime-only and must be explicitly reapplied after reopening when wanted.
///
/// OUTLINES: Group or ungroup row/column ranges, configure summary positions,
/// show a specific row/column outline level, inspect grouping state, or clear all groups.
/// </summary>
[ServiceCategory("sheet", "SheetStyle")]
[McpTool("worksheet_style", Title = "Worksheet Style Operations", Destructive = true, Category = "structure",
    Description = "Worksheet styling, visibility, protection, grouping, and outlines. OUTLINES: group/ungroup row or column ranges with axis Rows or Columns; get-outline-info reads level, hidden state, summary positions, and automatic styles; set-outline-settings accepts summaryRow above/below and summaryColumn left/right; show-outline-levels expands or collapses to row/column levels; clear-outline removes all groups. TAB COLORS: RGB values 0-255. VISIBILITY: visible, hidden, veryhidden. PROTECTION: set-protection accepts typed options replacing native permissions; omitted options use restrictive native defaults. get-protection reads all protected components, permissions and selection mode. userInterfaceOnly is runtime-only, not persisted after reopening. Filtering permission changes existing filters; sorting/deletion still require unlocked cells. Passwords are not returned. Use worksheet for lifecycle operations.")]
[McpReadOnlyActions("get-tab-color", "get-protection", "get-comment", "get-image-count", "get-shape-count",
    "get-page-setup", "get-page-breaks", "get-visibility", "get-outline-info")]
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
    /// When protecting, omitted options use Excel's restrictive native defaults.
    /// Supplied options replace the protection configuration, not patch existing permissions.
    /// UserInterfaceOnly is runtime-only; passwords are never returned.
    /// Throws exception on error.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="isProtected">Whether the worksheet should be protected</param>
    /// <param name="password">Optional password for protecting/unprotecting the sheet</param>
    /// <param name="options">Optional native protection permissions; valid only when protecting. Nested JSON uses camelCase.</param>
    [ServiceAction("set-protection")]
    OperationResult SetProtection(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] bool isProtected,
        string? password = null,
        SheetProtectionOptions? options = null);

    /// <summary>
    /// Reads protected components, all native permission flags, selection restriction, and runtime-only UI protection.
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
    /// <param name="orientation">Optional page orientation: 'portrait' or 'landscape'; omitted leaves unchanged</param>
    /// <param name="fitToPagesWide">Number of pages wide, or zero for unlimited; selects fit mode</param>
    /// <param name="fitToPagesTall">Number of pages tall, or zero for unlimited; selects fit mode</param>
    /// <param name="centerHorizontally">Whether to center the printout horizontally on the page</param>
    /// <param name="centerVertically">Whether to center the printout vertically on the page</param>
    /// <param name="pageSetupOptions">Native print scope, titles, point margins, headers/footers, paper, order and scaling. Nested camelCase; null preserves, empty text clears. Zoom conflicts with fit settings. Does not print or open preview.</param>
    [ServiceAction("set-page-setup")]
    OperationResult SetPageSetup(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        string? orientation = null,
        int? fitToPagesWide = null,
        int? fitToPagesTall = null,
        bool? centerHorizontally = null,
        bool? centerVertically = null,
        PageSetupOptions? pageSetupOptions = null);

    /// <summary>
    /// Reads worksheet page setup properties such as orientation and fit-to-page settings.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    [ServiceAction("get-page-setup")]
    SheetPageSetupResult GetPageSetup(IExcelBatch batch, [RequiredParameter] string sheetName);

    /// <summary>Reads every native horizontal/vertical page break in the worksheet's current print scope, including automatic/manual status and extent. Printer and scaling affect automatic breaks.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet to inspect</param>
    [ServiceAction("get-page-breaks")]
    SheetPageBreaksResult GetPageBreaks(IExcelBatch batch, [RequiredParameter] string sheetName);

    /// <summary>Replaces ALL manual page breaks on this worksheet; does not remove automatic breaks. Validates all positions before resetting. Empty lists explicitly clear. Excel allows at most 1026 breaks per axis.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet to change</param>
    /// <param name="pageBreakOptions">Required rows and columns lists with one-based positions before which to break. Both lists replace existing manual breaks, including those outside the current print scope.</param>
    [ServiceAction("set-page-breaks")]
    OperationResult SetPageBreaks(IExcelBatch batch, [RequiredParameter] string sheetName, [RequiredParameter] PageBreakOptions pageBreakOptions);

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

    /// <summary>
    /// Groups complete rows or columns covered by a range.
    /// Use row ranges such as '2:5' with axis Rows and column ranges such as 'B:D' with axis Columns.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Row or column range to group</param>
    /// <param name="axis">Grouping axis: Rows or Columns</param>
    [ServiceAction("group")]
    OperationResult Group(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string rangeAddress,
        [RequiredParameter]
        [FromString] OutlineAxis axis);

    /// <summary>
    /// Removes one grouping level from complete rows or columns covered by a range.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Grouped row or column range</param>
    /// <param name="axis">Grouping axis: Rows or Columns</param>
    [ServiceAction("ungroup")]
    OperationResult Ungroup(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string rangeAddress,
        [RequiredParameter]
        [FromString] OutlineAxis axis);

    /// <summary>
    /// Gets outline level, hidden state, summary positions, and automatic style settings.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Row or column range to inspect</param>
    /// <param name="axis">Outline axis: Rows or Columns</param>
    [ServiceAction("get-outline-info")]
    WorksheetOutlineResult GetOutlineInfo(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string rangeAddress,
        [RequiredParameter]
        [FromString] OutlineAxis axis);

    /// <summary>
    /// Sets worksheet outline summary positions or automatic styles.
    /// Omitted options remain unchanged.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="summaryRow">Summary row position: above or below</param>
    /// <param name="summaryColumn">Summary column position: left or right</param>
    /// <param name="automaticStyles">Whether Excel applies automatic outline styles</param>
    [ServiceAction("set-outline-settings")]
    OperationResult SetOutlineSettings(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        string? summaryRow = null,
        string? summaryColumn = null,
        bool? automaticStyles = null);

    /// <summary>
    /// Expands or collapses worksheet groups to the requested row and column outline levels.
    /// At least one level must be provided. Level values must be positive.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rowLevels">Optional row outline level to display</param>
    /// <param name="columnLevels">Optional column outline level to display</param>
    [ServiceAction("show-outline-levels")]
    OperationResult ShowOutlineLevels(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        int? rowLevels = null,
        int? columnLevels = null);

    /// <summary>
    /// Removes all row and column outline groups from a worksheet.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    [ServiceAction("clear-outline")]
    OperationResult ClearOutline(IExcelBatch batch, [RequiredParameter] string sheetName);
}
