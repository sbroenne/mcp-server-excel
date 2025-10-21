using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet management commands
/// </summary>
public interface ISheetCommands
{
    /// <summary>
    /// Lists all worksheets in the workbook
    /// </summary>
    WorksheetListResult List(string filePath);

    /// <summary>
    /// Reads data from a worksheet range
    /// </summary>
    WorksheetDataResult Read(string filePath, string sheetName, string range);

    /// <summary>
    /// Writes CSV data to a worksheet
    /// </summary>
    OperationResult Write(string filePath, string sheetName, string csvData);

    /// <summary>
    /// Creates a new worksheet
    /// </summary>
    OperationResult Create(string filePath, string sheetName);

    /// <summary>
    /// Renames a worksheet
    /// </summary>
    OperationResult Rename(string filePath, string oldName, string newName);

    /// <summary>
    /// Copies a worksheet
    /// </summary>
    OperationResult Copy(string filePath, string sourceName, string targetName);

    /// <summary>
    /// Deletes a worksheet
    /// </summary>
    OperationResult Delete(string filePath, string sheetName);

    /// <summary>
    /// Clears data from a worksheet range
    /// </summary>
    OperationResult Clear(string filePath, string sheetName, string range);

    /// <summary>
    /// Appends CSV data to a worksheet
    /// </summary>
    OperationResult Append(string filePath, string sheetName, string csvData);
}
