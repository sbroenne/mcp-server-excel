using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Session;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// VBA script management commands
/// </summary>
public interface IScriptCommands
{
    /// <summary>
    /// Lists all VBA modules and procedures in the workbook
    /// </summary>
    Task<ScriptListResult> ListAsync(IExcelBatch batch);

    /// <summary>
    /// Views VBA module code without exporting to file
    /// </summary>
    Task<ScriptViewResult> ViewAsync(IExcelBatch batch, string moduleName);

    /// <summary>
    /// Exports VBA module code to a file
    /// </summary>
    Task<OperationResult> ExportAsync(IExcelBatch batch, string moduleName, string outputFile);

    /// <summary>
    /// Imports VBA code from a file to create a new module
    /// </summary>
    Task<OperationResult> ImportAsync(IExcelBatch batch, string moduleName, string vbaFile);

    /// <summary>
    /// Updates an existing VBA module with new code
    /// </summary>
    Task<OperationResult> UpdateAsync(IExcelBatch batch, string moduleName, string vbaFile);

    /// <summary>
    /// Runs a VBA procedure with optional parameters
    /// </summary>
    Task<OperationResult> RunAsync(IExcelBatch batch, string procedureName, params string[] parameters);

    /// <summary>
    /// Deletes a VBA module
    /// </summary>
    Task<OperationResult> DeleteAsync(IExcelBatch batch, string moduleName);
}
