namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// File management commands for Excel workbooks
/// </summary>
public interface IFileCommands
{
    /// <summary>
    /// Creates an empty Excel workbook file
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx]</param>
    /// <returns>0 on success, 1 on error</returns>
    int CreateEmpty(string[] args);
}
