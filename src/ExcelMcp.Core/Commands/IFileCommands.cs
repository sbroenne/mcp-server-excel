using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// File management commands for Excel workbooks
/// </summary>
public interface IFileCommands
{
    /// <summary>
    /// Tests file existence, extension, lock state, read access, and deterministic
    /// IRM/AIP read-only and visible-session requirements without inspecting
    /// workbook structure. Structural validity and openability remain false until
    /// Excel opens the workbook.
    /// </summary>
    /// <param name="filePath">Path to the Excel file to validate</param>
    /// <returns>Canonical file metadata shared by CLI and MCP</returns>
    FileValidationInfo Test(string filePath);
}
