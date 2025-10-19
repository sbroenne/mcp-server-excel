namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// VBA script management commands
/// </summary>
public interface IScriptCommands
{
    /// <summary>
    /// Lists all VBA modules and procedures in the workbook
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsm]</param>
    /// <returns>0 on success, 1 on error</returns>
    int List(string[] args);
    
    /// <summary>
    /// Exports VBA module code to a file
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsm, moduleName, outputFile]</param>
    /// <returns>0 on success, 1 on error</returns>
    int Export(string[] args);
    
    /// <summary>
    /// Imports VBA code from a file to create a new module
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsm, moduleName, vbaFile]</param>
    /// <returns>0 on success, 1 on error</returns>
    Task<int> Import(string[] args);
    
    /// <summary>
    /// Updates an existing VBA module with new code
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsm, moduleName, vbaFile]</param>
    /// <returns>0 on success, 1 on error</returns>
    Task<int> Update(string[] args);
    
    /// <summary>
    /// Runs a VBA procedure with optional parameters
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsm, module.procedure, param1, param2, ...]</param>
    /// <returns>0 on success, 1 on error</returns>
    int Run(string[] args);
}
