using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Session;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet management commands - all operations are batch-aware for performance.
/// Use ExcelSession.BeginBatchAsync() to create a batch, then pass it to these methods.
/// </summary>
public interface ISheetCommands
{
    /// <summary>
    /// Lists all worksheets in the workbook
    /// </summary>
    Task<WorksheetListResult> ListAsync(IExcelBatch batch);

    /// <summary>
    /// Reads data from a worksheet range
    /// </summary>
    Task<WorksheetDataResult> ReadAsync(IExcelBatch batch, string sheetName, string? range = null);

    /// <summary>
    /// Writes CSV data to a worksheet
    /// </summary>
    Task<OperationResult> WriteAsync(IExcelBatch batch, string sheetName, string csvData);

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

    /// <summary>
    /// Clears data from a worksheet range
    /// </summary>
    Task<OperationResult> ClearAsync(IExcelBatch batch, string sheetName, string? range = null);

    /// <summary>
    /// Appends CSV data to a worksheet
    /// </summary>
    Task<OperationResult> AppendAsync(IExcelBatch batch, string sheetName, string csvData);
}
