namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>
/// Utilities for path validation and normalization in CLI commands.
/// </summary>
internal static class PathValidator
{
    private static readonly string[] ValidExcelExtensions = [".xlsx", ".xlsm", ".xlsb", ".xls"];

    /// <summary>
    /// Validates that a file path is a valid Excel file path.
    /// </summary>
    /// <param name="filePath">The file path to validate</param>
    /// <param name="mustExist">If true, file must exist</param>
    /// <param name="errorMessage">Error message if validation fails</param>
    /// <returns>True if valid, false otherwise</returns>
    public static bool ValidateExcelPath(string? filePath, bool mustExist, out string? errorMessage)
    {
        errorMessage = null;

        if (string.IsNullOrWhiteSpace(filePath))
        {
            errorMessage = "File path is required.";
            return false;
        }

        // Normalize path
        var normalizedPath = NormalizePath(filePath);

        // Check extension
        var extension = Path.GetExtension(normalizedPath).ToLowerInvariant();
        if (!ValidExcelExtensions.Contains(extension))
        {
            errorMessage = $"Invalid file extension '{extension}'. Valid extensions: {string.Join(", ", ValidExcelExtensions)}";
            return false;
        }

        // Check if file exists (if required)
        if (mustExist && !File.Exists(normalizedPath))
        {
            errorMessage = $"File not found: {normalizedPath}";
            return false;
        }

        // Check parent directory exists (for new files)
        if (!mustExist)
        {
            var directory = Path.GetDirectoryName(normalizedPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                errorMessage = $"Directory not found: {directory}";
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// Validates that a directory path exists.
    /// </summary>
    public static bool ValidateDirectory(string? directoryPath, out string? errorMessage)
    {
        errorMessage = null;

        if (string.IsNullOrWhiteSpace(directoryPath))
        {
            errorMessage = "Directory path is required.";
            return false;
        }

        var normalizedPath = NormalizePath(directoryPath);

        if (!Directory.Exists(normalizedPath))
        {
            errorMessage = $"Directory not found: {normalizedPath}";
            return false;
        }

        return true;
    }

    /// <summary>
    /// Validates and normalizes a file path, returning the full path.
    /// </summary>
    public static string? GetNormalizedPath(string? filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            return null;

        return NormalizePath(filePath);
    }

    /// <summary>
    /// Normalizes a file path to full path format.
    /// </summary>
    public static string NormalizePath(string filePath)
    {
        // Handle relative paths
        if (!Path.IsPathRooted(filePath))
        {
            filePath = Path.Combine(Environment.CurrentDirectory, filePath);
        }

        // Normalize to full path (resolves . and ..)
        return Path.GetFullPath(filePath);
    }

    /// <summary>
    /// Checks if a path is within an allowed directory (security check).
    /// </summary>
    public static bool IsWithinAllowedDirectory(string filePath, string allowedDirectory)
    {
        var normalizedFile = NormalizePath(filePath);
        var normalizedAllowed = NormalizePath(allowedDirectory);

        return normalizedFile.StartsWith(normalizedAllowed, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Checks if a path contains potentially dangerous characters.
    /// </summary>
    public static bool HasInvalidPathCharacters(string path)
    {
        var invalidChars = Path.GetInvalidPathChars();
        return path.Any(c => invalidChars.Contains(c));
    }

    /// <summary>
    /// Gets a unique file path by adding a number suffix if file exists.
    /// </summary>
    public static string GetUniqueFilePath(string filePath)
    {
        if (!File.Exists(filePath))
            return filePath;

        var directory = Path.GetDirectoryName(filePath) ?? ".";
        var nameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
        var extension = Path.GetExtension(filePath);

        var counter = 1;
        string newPath;
        do
        {
            newPath = Path.Combine(directory, $"{nameWithoutExt}_{counter}{extension}");
            counter++;
        } while (File.Exists(newPath) && counter < 1000);

        return newPath;
    }
}
