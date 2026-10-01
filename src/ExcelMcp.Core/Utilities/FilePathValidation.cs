namespace Sbroenne.ExcelMcp.Core.Utilities;

/// <summary>
/// Shared validation for public workbook path inputs.
/// </summary>
public static class FilePathValidation
{
    /// <summary>
    /// Requires and normalizes an absolute file path for the current platform.
    /// </summary>
    public static string NormalizeAbsolutePath(string filePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        if (!Path.IsPathFullyQualified(filePath))
        {
            throw new ArgumentException(
                $"File path must be an absolute path: '{filePath}'.",
                nameof(filePath));
        }

        return Path.GetFullPath(filePath);
    }
}
