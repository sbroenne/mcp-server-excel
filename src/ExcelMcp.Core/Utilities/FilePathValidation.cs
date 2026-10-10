using Sbroenne.ExcelMcp.ComInterop;

namespace Sbroenne.ExcelMcp.Core.Utilities;

/// <summary>
/// Shared validation for public workbook path inputs.
/// </summary>
public static class FilePathValidation
{
    /// <summary>
    /// Normalizes an existing workbook's Windows path or direct SharePoint HTTPS URL.
    /// </summary>
    public static string NormalizeWorkbookLocation(string location) => WorkbookLocation.Normalize(location);

    /// <summary>
    /// Identifies a normalized remote workbook location.
    /// </summary>
    public static bool IsRemoteWorkbook(string location) => WorkbookLocation.IsRemote(location);

    /// <summary>
    /// Requires and normalizes an absolute Windows file path.
    /// </summary>
    public static string NormalizeAbsoluteWindowsPath(string filePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        if (!Path.IsPathFullyQualified(filePath))
        {
            throw new ArgumentException(
                $"File path must be an absolute Windows path: '{filePath}'.",
                nameof(filePath));
        }

        return Path.GetFullPath(filePath);
    }
}
