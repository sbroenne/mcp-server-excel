namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Named range/parameter management commands
/// </summary>
public interface IParameterCommands
{
    /// <summary>
    /// Lists all named ranges in the workbook
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx]</param>
    /// <returns>0 on success, 1 on error</returns>
    int List(string[] args);
    
    /// <summary>
    /// Sets the value of a named range
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, paramName, value]</param>
    /// <returns>0 on success, 1 on error</returns>
    int Set(string[] args);
    
    /// <summary>
    /// Gets the value of a named range
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, paramName]</param>
    /// <returns>0 on success, 1 on error</returns>
    int Get(string[] args);
    
    /// <summary>
    /// Creates a new named range
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, paramName, reference]</param>
    /// <returns>0 on success, 1 on error</returns>
    int Create(string[] args);
    
    /// <summary>
    /// Deletes a named range
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, paramName]</param>
    /// <returns>0 on success, 1 on error</returns>
    int Delete(string[] args);
}
