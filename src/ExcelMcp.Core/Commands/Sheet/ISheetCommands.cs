using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet lifecycle management: create, rename, copy, delete, move, list sheets.
/// Use range for data operations. Use sheetstyle for tab colors and visibility.
///
/// ATOMIC OPERATIONS: 'copy-to-file' and 'move-to-file' don't require a session -
/// they open/close files automatically.
///
/// POSITIONING: For 'move', 'copy-to-file', 'move-to-file' - use 'before' OR 'after'
/// (not both) to position the sheet relative to another. If neither specified, moves to end.
/// 
/// NOTE: MCP tool is manually implemented in ExcelWorksheetTool.cs to properly handle
/// mixed session requirements (copy-to-file and move-to-file are atomic and don't need sessions).
/// </summary>
[ServiceCategory("sheet", "Sheet")]
public interface ISheetCommands
{
    // === LIFECYCLE OPERATIONS ===

    /// <summary>
    /// Lists all worksheets in the workbook.
    /// For multi-workbook batches, specify filePath to list sheets from a specific workbook.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="filePath">Optional file path when batch contains multiple workbooks. If omitted, uses primary workbook.</param>
    [ServiceAction("list")]
    WorksheetListResult List(IExcelBatch batch, string? filePath = null);

    /// <summary>
    /// Creates a new worksheet.
    /// For multi-workbook batches, specify filePath to create in a specific workbook.
    /// Throws exception on error.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name for the new worksheet</param>
    /// <param name="filePath">Optional file path when batch contains multiple workbooks. If omitted, creates in primary workbook.</param>
    [ServiceAction("create")]
    OperationResult Create(IExcelBatch batch, [RequiredParameter] string sheetName, string? filePath = null);

    /// <summary>
    /// Renames a worksheet.
    /// Throws exception on error.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="oldName">Current name of the worksheet</param>
    /// <param name="newName">New name for the worksheet</param>
    [ServiceAction("rename")]
    OperationResult Rename(IExcelBatch batch, [RequiredParameter] string oldName, [RequiredParameter] string newName);

    /// <summary>
    /// Copies a worksheet.
    /// Throws exception on error.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sourceName">Name of the source worksheet</param>
    /// <param name="targetName">Name for the copied worksheet</param>
    [ServiceAction("copy")]
    OperationResult Copy(IExcelBatch batch, [RequiredParameter] string sourceName, [RequiredParameter] string targetName);

    /// <summary>
    /// Deletes a worksheet.
    /// Throws exception on error.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet to delete</param>
    [ServiceAction("delete")]
    OperationResult Delete(IExcelBatch batch, [RequiredParameter] string sheetName);

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
    [ServiceAction("move")]
    OperationResult Move(IExcelBatch batch, [RequiredParameter] string sheetName, string? beforeSheet = null, string? afterSheet = null);

    // === ATOMIC CROSS-FILE OPERATIONS ===
    // These operations don't require a session - they create temporary Excel instances internally.

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
    [ServiceAction("copy-to-file")]
    OperationResult CopyToFile(
        [RequiredParameter] string sourceFile,
        [RequiredParameter] string sourceSheet,
        [RequiredParameter] string targetFile,
        string? targetSheetName = null,
        string? beforeSheet = null,
        string? afterSheet = null);

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
    [ServiceAction("move-to-file")]
    OperationResult MoveToFile(
        [RequiredParameter] string sourceFile,
        [RequiredParameter] string sourceSheet,
        [RequiredParameter] string targetFile,
        string? beforeSheet = null,
        string? afterSheet = null);
}



