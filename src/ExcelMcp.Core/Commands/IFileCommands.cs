using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// File management commands for Excel workbooks
/// </summary>
public interface IFileCommands
{
    /// <summary>
    /// Tests if a file exists and is a valid Excel file
    /// </summary>
    /// <param name="filePath">Path to the Excel file to validate</param>
    /// <returns>File validation details including existence, size, extension, and validity information</returns>
    FileValidationInfo Test(string filePath);
}

