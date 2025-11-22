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
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name for the new worksheet</param>
    /// <param name="filePath">Optional file path when batch contains multiple workbooks. If omitted, creates in primary workbook.</param>
    OperationResult Create(IExcelBatch batch, string sheetName, string? filePath = null);

    /// <summary>
    /// Renames a worksheet
    /// </summary>
    OperationResult Rename(IExcelBatch batch, string oldName, string newName);

    /// <summary>
    /// Copies a worksheet
    /// </summary>
    OperationResult Copy(IExcelBatch batch, string sourceName, string targetName);

    /// <summary>
    /// Deletes a worksheet
    /// </summary>
    OperationResult Delete(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Moves a worksheet to a new position within the workbook.
    /// Use either beforeSheet OR afterSheet to specify position (not both).
    /// If neither is specified, sheet moves to the end.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the sheet to move</param>
    /// <param name="beforeSheet">Optional: Name of sheet to position before</param>
    /// <param name="afterSheet">Optional: Name of sheet to position after</param>
    OperationResult Move(IExcelBatch batch, string sheetName, string? beforeSheet = null, string? afterSheet = null);

    /// <summary>
    /// Copies a worksheet to another workbook.
    /// Both workbooks must be open in the same batch (multi-workbook batch).
    /// </summary>
    /// <param name="batch">Excel batch containing both workbooks</param>
    /// <param name="sourceFile">Source workbook file path</param>
    /// <param name="sourceSheet">Name of sheet to copy</param>
    /// <param name="targetFile">Target workbook file path</param>
    /// <param name="targetSheetName">Optional: New name for the copied sheet in target workbook</param>
    /// <param name="beforeSheet">Optional: Name of sheet in target workbook to position before</param>
    /// <param name="afterSheet">Optional: Name of sheet in target workbook to position after</param>
    OperationResult CopyToWorkbook(IExcelBatch batch, string sourceFile, string sourceSheet, string targetFile, string? targetSheetName = null, string? beforeSheet = null, string? afterSheet = null);

    /// <summary>
    /// Moves a worksheet to another workbook.
    /// Both workbooks must be open in the same batch (multi-workbook batch).
    /// The sheet will be removed from the source workbook.
    /// </summary>
    /// <param name="batch">Excel batch containing both workbooks</param>
    /// <param name="sourceFile">Source workbook file path</param>
    /// <param name="sourceSheet">Name of sheet to move</param>
    /// <param name="targetFile">Target workbook file path</param>
    /// <param name="beforeSheet">Optional: Name of sheet in target workbook to position before</param>
    /// <param name="afterSheet">Optional: Name of sheet in target workbook to position after</param>
    OperationResult MoveToWorkbook(IExcelBatch batch, string sourceFile, string sourceSheet, string targetFile, string? beforeSheet = null, string? afterSheet = null);

    // === TAB COLOR OPERATIONS ===

    /// <summary>
    /// Sets the tab color for a worksheet using RGB values (0-255 each).
    /// Excel uses BGR format internally, conversion is handled automatically.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="red">Red component (0-255)</param>
    /// <param name="green">Green component (0-255)</param>
    /// <param name="blue">Blue component (0-255)</param>
    OperationResult SetTabColor(IExcelBatch batch, string sheetName, int red, int green, int blue);

    /// <summary>
    /// Gets the tab color for a worksheet.
    /// Returns RGB values and hex color, or HasColor=false if no color is set.
    /// </summary>
    TabColorResult GetTabColor(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Clears the tab color for a worksheet (resets to default)
    /// </summary>
    OperationResult ClearTabColor(IExcelBatch batch, string sheetName);

    // === VISIBILITY OPERATIONS ===

    /// <summary>
    /// Sets worksheet visibility level.
    /// - Visible: Normal visible state
    /// - Hidden: Hidden via UI, user can unhide
    /// - VeryHidden: Requires code to unhide (security/protection)
    /// </summary>
    OperationResult SetVisibility(IExcelBatch batch, string sheetName, SheetVisibility visibility);

    /// <summary>
    /// Gets worksheet visibility level
    /// </summary>
    SheetVisibilityResult GetVisibility(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Shows a hidden or very hidden worksheet.
    /// Convenience method equivalent to SetVisibilityAsync(..., SheetVisibility.Visible)
    /// </summary>
    OperationResult Show(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Hides a worksheet (user can unhide via Excel UI).
    /// Convenience method equivalent to SetVisibilityAsync(..., SheetVisibility.Hidden)
    /// </summary>
    OperationResult Hide(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Very hides a worksheet (requires code to unhide, for protection).
    /// Convenience method equivalent to SetVisibilityAsync(..., SheetVisibility.VeryHidden)
    /// </summary>
    OperationResult VeryHide(IExcelBatch batch, string sheetName);
}

