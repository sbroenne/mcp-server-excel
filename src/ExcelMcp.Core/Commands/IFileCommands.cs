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
    /// Validates an Excel file
    /// </summary>
    /// <param name="filePath">Path to the Excel file to validate</param>
    /// <returns>File validation result</returns>
    FileValidationResult Validate(string filePath);
}
