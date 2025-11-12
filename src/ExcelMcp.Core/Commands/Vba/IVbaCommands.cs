using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// VBA script management commands
/// </summary>
public interface IVbaCommands
{
    /// <summary>
    /// Lists all VBA modules and procedures in the workbook
    /// </summary>
    Task<VbaListResult> ListAsync(IExcelBatch batch);

    /// <summary>
    /// Lists all VBA modules and procedures in the workbook (filePath-based API)
    /// </summary>
    Task<VbaListResult> ListAsync(string filePath);

    /// <summary>
    /// Views VBA module code without exporting to file
    /// </summary>
    Task<VbaViewResult> ViewAsync(IExcelBatch batch, string moduleName);

    /// <summary>
    /// Views VBA module code without exporting to file (filePath-based API)
    /// </summary>
    Task<VbaViewResult> ViewAsync(string filePath, string moduleName);

    /// <summary>
    /// Exports VBA module code to a file
    /// </summary>
    Task<OperationResult> ExportAsync(IExcelBatch batch, string moduleName, string outputFile);

    /// <summary>
    /// Exports VBA module code to a file (filePath-based API)
    /// </summary>
    Task<OperationResult> ExportAsync(string filePath, string moduleName, string outputFile);

    /// <summary>
    /// Imports VBA code from a file to create a new module
    /// </summary>
    Task<OperationResult> ImportAsync(IExcelBatch batch, string moduleName, string vbaFile);

    /// <summary>
    /// Imports VBA code from a file to create a new module (filePath-based API)
    /// </summary>
    Task<OperationResult> ImportAsync(string filePath, string moduleName, string vbaFile);

    /// <summary>
    /// Updates an existing VBA module with new code
    /// </summary>
    Task<OperationResult> UpdateAsync(IExcelBatch batch, string moduleName, string vbaFile);

    /// <summary>
    /// Updates an existing VBA module with new code (filePath-based API)
    /// </summary>
    Task<OperationResult> UpdateAsync(string filePath, string moduleName, string vbaFile);

    /// <summary>
    /// Runs a VBA procedure with optional parameters
    /// </summary>
    Task<OperationResult> RunAsync(IExcelBatch batch, string procedureName, TimeSpan? timeout, params string[] parameters);

    /// <summary>
    /// Runs a VBA procedure with optional parameters (filePath-based API)
    /// </summary>
    Task<OperationResult> RunAsync(string filePath, string procedureName, TimeSpan? timeout, params string[] parameters);

    /// <summary>
    /// Deletes a VBA module
    /// </summary>
    Task<OperationResult> DeleteAsync(IExcelBatch batch, string moduleName);

    /// <summary>
    /// Deletes a VBA module (filePath-based API)
    /// </summary>
    Task<OperationResult> DeleteAsync(string filePath, string moduleName);
}
