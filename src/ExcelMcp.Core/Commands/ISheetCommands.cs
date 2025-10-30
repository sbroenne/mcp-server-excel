using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet lifecycle management commands - create, rename, copy, delete worksheets.
/// Data operations (read, write, clear) moved to IRangeCommands for unified range API.
/// All operations are batch-aware for performance.
/// Use ExcelSession.BeginBatchAsync() to create a batch, then pass it to these methods.
/// </summary>
public interface ISheetCommands
{
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
}
