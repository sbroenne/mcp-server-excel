namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query management commands
/// </summary>
public interface IPowerQueryCommands
{
    /// <summary>
    /// Lists all Power Query queries in the workbook
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx]</param>
    /// <returns>0 on success, 1 on error</returns>
    int List(string[] args);
    
    /// <summary>
    /// Views the M code of a Power Query
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, queryName]</param>
    /// <returns>0 on success, 1 on error</returns>
    int View(string[] args);
    
    /// <summary>
    /// Updates an existing Power Query with new M code
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, queryName, mCodeFile]</param>
    /// <returns>0 on success, 1 on error</returns>
    Task<int> Update(string[] args);
    
    /// <summary>
    /// Exports a Power Query's M code to a file
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, queryName, outputFile]</param>
    /// <returns>0 on success, 1 on error</returns>
    Task<int> Export(string[] args);
    
    /// <summary>
    /// Imports M code from a file to create a new Power Query
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, queryName, mCodeFile]</param>
    /// <returns>0 on success, 1 on error</returns>
    Task<int> Import(string[] args);
    
    /// <summary>
    /// Refreshes a Power Query to update its data
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, queryName]</param>
    /// <returns>0 on success, 1 on error</returns>
    int Refresh(string[] args);
    
    /// <summary>
    /// Shows errors from Power Query operations
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, queryName]</param>
    /// <returns>0 on success, 1 on error</returns>
    int Errors(string[] args);
    
    /// <summary>
    /// Loads a connection-only Power Query to a worksheet
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, queryName, sheetName]</param>
    /// <returns>0 on success, 1 on error</returns>
    int LoadTo(string[] args);
    
    /// <summary>
    /// Deletes a Power Query from the workbook
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, queryName]</param>
    /// <returns>0 on success, 1 on error</returns>
    int Delete(string[] args);
}
