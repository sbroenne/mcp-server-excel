using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// VBA script management commands
/// </summary>
public interface IScriptCommands
{
    /// <summary>
    /// Lists all VBA modules and procedures in the workbook
    /// </summary>
    ScriptListResult List(string filePath);
    
    /// <summary>
    /// Exports VBA module code to a file
    /// </summary>
    Task<OperationResult> Export(string filePath, string moduleName, string outputFile);
    
    /// <summary>
    /// Imports VBA code from a file to create a new module
    /// </summary>
    Task<OperationResult> Import(string filePath, string moduleName, string vbaFile);
    
    /// <summary>
    /// Updates an existing VBA module with new code
    /// </summary>
    Task<OperationResult> Update(string filePath, string moduleName, string vbaFile);
    
    /// <summary>
    /// Runs a VBA procedure with optional parameters
    /// </summary>
    OperationResult Run(string filePath, string procedureName, params string[] parameters);
    
    /// <summary>
    /// Deletes a VBA module
    /// </summary>
    OperationResult Delete(string filePath, string moduleName);
}
