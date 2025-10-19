namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet management commands
/// </summary>
public interface ISheetCommands
{
    /// <summary>
    /// Lists all worksheets in the workbook
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx]</param>
    /// <returns>0 on success, 1 on error</returns>
    int List(string[] args);
    
    /// <summary>
    /// Reads data from a worksheet range
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, sheetName, range]</param>
    /// <returns>0 on success, 1 on error</returns>
    int Read(string[] args);
    
    /// <summary>
    /// Writes CSV data to a worksheet
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, sheetName, csvFile]</param>
    /// <returns>0 on success, 1 on error</returns>
    Task<int> Write(string[] args);
    
    /// <summary>
    /// Copies a worksheet within the workbook
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, sourceSheet, targetSheet]</param>
    /// <returns>0 on success, 1 on error</returns>
    int Copy(string[] args);
    
    /// <summary>
    /// Deletes a worksheet from the workbook
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, sheetName]</param>
    /// <returns>0 on success, 1 on error</returns>
    int Delete(string[] args);
    
    /// <summary>
    /// Creates a new worksheet in the workbook
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, sheetName]</param>
    /// <returns>0 on success, 1 on error</returns>
    int Create(string[] args);
    
    /// <summary>
    /// Renames an existing worksheet
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, oldName, newName]</param>
    /// <returns>0 on success, 1 on error</returns>
    int Rename(string[] args);
    
    /// <summary>
    /// Clears data from a worksheet range
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, sheetName, range]</param>
    /// <returns>0 on success, 1 on error</returns>
    int Clear(string[] args);
    
    /// <summary>
    /// Appends CSV data to existing worksheet content
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, sheetName, csvFile]</param>
    /// <returns>0 on success, 1 on error</returns>
    int Append(string[] args);
}
