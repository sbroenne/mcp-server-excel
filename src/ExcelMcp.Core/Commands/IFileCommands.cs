using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// File management commands for Excel workbooks
/// </summary>
public interface IFileCommands
{
    /// <summary>
    /// Checks local file details and reports IRM/AIP protection before a read-only Excel open.
    /// Excel determines editing permissions after authentication.
    /// URLs report unavailable local file details without opening Excel.
    /// </summary>
    /// <param name="filePath">Absolute Windows path or URL</param>
    /// <returns>Canonical file metadata shared by CLI and MCP</returns>
    FileValidationInfo Test(string filePath);
}
