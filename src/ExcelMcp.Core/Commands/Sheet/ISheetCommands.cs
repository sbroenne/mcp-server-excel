using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet lifecycle and appearance management commands.
/// - Lifecycle: create, rename, copy, delete worksheets
/// - Appearance: tab colors, visibility levels
/// Data operations (read, write, clear) moved to IRangeCommands for unified range API.
/// All operations are batch-aware for performance.
/// Use ExcelSession.BeginBatchAsync() to create a batch, then pass it to these methods.
/// </summary>
public interface ISheetCommands
{
    // === LIFECYCLE OPERATIONS ===

    /// <summary>
    /// Lists all worksheets in the workbook
    /// </summary>
    Task<WorksheetListResult> ListAsync(IExcelBatch batch);

    /// <summary>
    /// Creates a new worksheet
    /// </summary>
    Task<OperationResult> CreateAsync(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Renames a worksheet
    /// </summary>
    Task<OperationResult> RenameAsync(IExcelBatch batch, string oldName, string newName);

    /// <summary>
    /// Copies a worksheet
    /// </summary>
    Task<OperationResult> CopyAsync(IExcelBatch batch, string sourceName, string targetName);

    /// <summary>
    /// Deletes a worksheet
    /// </summary>
    Task<OperationResult> DeleteAsync(IExcelBatch batch, string sheetName);

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
    Task<OperationResult> SetTabColorAsync(IExcelBatch batch, string sheetName, int red, int green, int blue);

    /// <summary>
    /// Gets the tab color for a worksheet.
    /// Returns RGB values and hex color, or HasColor=false if no color is set.
    /// </summary>
    Task<TabColorResult> GetTabColorAsync(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Clears the tab color for a worksheet (resets to default)
    /// </summary>
    Task<OperationResult> ClearTabColorAsync(IExcelBatch batch, string sheetName);

    // === VISIBILITY OPERATIONS ===

    /// <summary>
    /// Sets worksheet visibility level.
    /// - Visible: Normal visible state
    /// - Hidden: Hidden via UI, user can unhide
    /// - VeryHidden: Requires code to unhide (security/protection)
    /// </summary>
    Task<OperationResult> SetVisibilityAsync(IExcelBatch batch, string sheetName, SheetVisibility visibility);

    /// <summary>
    /// Gets worksheet visibility level
    /// </summary>
    Task<SheetVisibilityResult> GetVisibilityAsync(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Shows a hidden or very hidden worksheet.
    /// Convenience method equivalent to SetVisibilityAsync(..., SheetVisibility.Visible)
    /// </summary>
    Task<OperationResult> ShowAsync(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Hides a worksheet (user can unhide via Excel UI).
    /// Convenience method equivalent to SetVisibilityAsync(..., SheetVisibility.Hidden)
    /// </summary>
    Task<OperationResult> HideAsync(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Very hides a worksheet (requires code to unhide, for protection).
    /// Convenience method equivalent to SetVisibilityAsync(..., SheetVisibility.VeryHidden)
    /// </summary>
    Task<OperationResult> VeryHideAsync(IExcelBatch batch, string sheetName);
}
