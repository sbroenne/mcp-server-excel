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
    /// <exception cref="ArgumentException">File path invalid or extension not .xlsx/.xlsm</exception>
    /// <exception cref="InvalidOperationException">Directory creation failed</exception>
    void CreateEmpty(string filePath, bool overwriteIfExists = false);

    /// <summary>
    /// Tests if a file exists and is a valid Excel file
    /// </summary>
    /// <param name="filePath">Path to the Excel file to validate</param>
    /// <returns>File validation details including existence, size, extension, and validity information</returns>
    FileValidationInfo Test(string filePath);
}

