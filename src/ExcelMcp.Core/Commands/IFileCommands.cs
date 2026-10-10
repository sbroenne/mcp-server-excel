using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// File management commands for Excel workbooks
/// </summary>
public interface IFileCommands
{
    /// <summary>
    /// Tests file existence, Excel extension validity (.xlsx, .xlsm, .xlsb, .xls), file access, and deterministic
    /// IRM/AIP visible-session requirements before Service open validation.
    /// Excel determines editing permissions after authentication.
    /// Direct SharePoint HTTPS URLs report a visible-authentication requirement without a local file probe.
    /// </summary>
    /// <param name="filePath">Absolute Windows path or direct SharePoint/OneDrive for Business HTTPS workbook URL</param>
    /// <returns>Canonical file metadata shared by CLI and MCP</returns>
    FileValidationInfo Test(string filePath);
}
