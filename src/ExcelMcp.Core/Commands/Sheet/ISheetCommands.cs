using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet lifecycle and appearance management commands.
/// - Lifecycle: create, rename, copy, delete worksheets
/// - Appearance: tab colors, visibility levels
/// Data operations (read, write, clear) moved to IRangeCommands for unified range API.
/// All operations are batch-aware for performance.
/// Use ExcelSession.BeginBatch() to create a batch, then pass it to these methods.
/// </summary>
public interface ISheetCommands
{
    // === LIFECYCLE OPERATIONS ===

    /// <summary>
    /// Lists all worksheets in the workbook.
    /// For multi-workbook batches, specify filePath to list sheets from a specific workbook.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="filePath">Optional file path when batch contains multiple workbooks. If omitted, uses primary workbook.</param>
    WorksheetListResult List(IExcelBatch batch, string? filePath = null);

    /// <summary>
    /// Creates a new worksheet.
    /// For multi-workbook batches, specify filePath to create in a specific workbook.
    /// Throws exception on error.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name for the new worksheet</param>
    /// <param name="filePath">Optional file path when batch contains multiple workbooks. If omitted, creates in primary workbook.</param>
    void Create(IExcelBatch batch, string sheetName, string? filePath = null);

    /// <summary>
    /// Renames a worksheet.
    /// Throws exception on error.
    /// </summary>
    void Rename(IExcelBatch batch, string oldName, string newName);

    /// <summary>
    /// Copies a worksheet.
    /// Throws exception on error.
    /// </summary>
    void Copy(IExcelBatch batch, string sourceName, string targetName);

    /// <summary>
    /// Deletes a worksheet.
    /// Throws exception on error.
    /// </summary>
    void Delete(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Moves a worksheet to a new position within the workbook.
    /// Use either beforeSheet OR afterSheet to specify position (not both).
    /// If neither is specified, sheet moves to the end.
    /// Throws exception on error.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the sheet to move</param>
    /// <param name="beforeSheet">Optional: Name of sheet to position before</param>
    /// <param name="afterSheet">Optional: Name of sheet to position after</param>
    void Move(IExcelBatch batch, string sheetName, string? beforeSheet = null, string? afterSheet = null);

    // === ATOMIC CROSS-FILE OPERATIONS ===

    /// <summary>
    /// Copies a worksheet to another file (atomic operation - no session required).
    /// Creates a temporary Excel instance, opens both files, performs the copy,
    /// saves the target file, and closes both files.
    /// </summary>
    /// <param name="sourceFile">Full path to the source workbook</param>
    /// <param name="sourceSheet">Name of the sheet to copy</param>
    /// <param name="targetFile">Full path to the target workbook</param>
    /// <param name="targetSheetName">Optional: New name for the copied sheet (default: keeps original name)</param>
    /// <param name="beforeSheet">Optional: Position before this sheet in target</param>
    /// <param name="afterSheet">Optional: Position after this sheet in target</param>
    void CopyToFile(string sourceFile, string sourceSheet, string targetFile, string? targetSheetName = null, string? beforeSheet = null, string? afterSheet = null);

    /// <summary>
    /// Moves a worksheet to another file (atomic operation - no session required).
    /// Creates a temporary Excel instance, opens both files, performs the move,
    /// saves both files, and closes them.
    /// This is the RECOMMENDED way to move sheets between files.
    /// </summary>
    /// <param name="sourceFile">Full path to the source workbook</param>
    /// <param name="sourceSheet">Name of the sheet to move</param>
    /// <param name="targetFile">Full path to the target workbook</param>
    /// <param name="beforeSheet">Optional: Position before this sheet in target</param>
    /// <param name="afterSheet">Optional: Position after this sheet in target</param>
    void MoveToFile(string sourceFile, string sourceSheet, string targetFile, string? beforeSheet = null, string? afterSheet = null);

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
    void SetTabColor(IExcelBatch batch, string sheetName, int red, int green, int blue);

    /// <summary>
    /// Gets the tab color for a worksheet.
    /// Returns RGB values and hex color, or HasColor=false if no color is set.
    /// </summary>
    TabColorResult GetTabColor(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Clears the tab color for a worksheet (resets to default).
    /// Throws exception on error.
    /// </summary>
    void ClearTabColor(IExcelBatch batch, string sheetName);

    // === VISIBILITY OPERATIONS ===

    /// <summary>
    /// Sets worksheet visibility level.
    /// - Visible: Normal visible state
    /// - Hidden: Hidden via UI, user can unhide
    /// - VeryHidden: Requires code to unhide (security/protection)
    /// Throws exception on error.
    /// </summary>
    void SetVisibility(IExcelBatch batch, string sheetName, SheetVisibility visibility);

    /// <summary>
    /// Gets worksheet visibility level
    /// </summary>
    SheetVisibilityResult GetVisibility(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Shows a hidden or very hidden worksheet.
    /// Convenience method equivalent to SetVisibilityAsync(..., SheetVisibility.Visible).
    /// Throws exception on error.
    /// </summary>
    void Show(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Hides a worksheet (user can unhide via Excel UI).
    /// Convenience method equivalent to SetVisibilityAsync(..., SheetVisibility.Hidden).
    /// Throws exception on error.
    /// </summary>
    void Hide(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Very hides a worksheet (requires code to unhide, for protection).
    /// Convenience method equivalent to SetVisibilityAsync(..., SheetVisibility.VeryHidden).
    /// Throws exception on error.
    /// </summary>
    void VeryHide(IExcelBatch batch, string sheetName);
}

