using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// File management commands for Excel workbooks
/// </summary>
public interface IFileCommands
{
    /// <summary>
    /// Creates an empty Excel workbook file
    /// </summary>
    /// <param name="filePath">Path to the Excel file to create</param>
    /// <param name="overwriteIfExists">Whether to overwrite if file already exists</param>
    /// <returns>Operation result</returns>
    OperationResult CreateEmpty(string filePath, bool overwriteIfExists = false);

    /// <summary>
    /// Tests if a file exists and is a valid Excel file
    /// </summary>
    /// <param name="filePath">Path to the Excel file to validate</param>
    /// <returns>File validation result with existence, size, extension, and validity information</returns>
    FileValidationResult Test(string filePath);
}

