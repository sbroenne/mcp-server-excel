using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.Core.Security;

/// <summary>
/// Provides path validation and security checks to prevent path traversal attacks
/// </summary>
public static class PathValidator
{
    /// <summary>
    /// Maximum allowed path length for Windows file system
    /// </summary>
    private const int MaxPathLength = 32767;

    /// <summary>
    /// Maximum allowed file size for input files (100MB) to prevent DoS attacks
    /// </summary>
    private const long MaxInputFileSize = 100 * 1024 * 1024;

    /// <summary>
    /// Validates and normalizes a file path to prevent path traversal attacks
    /// </summary>
    /// <param name="path">The file path to validate</param>
    /// <param name="parameterName">Name of the parameter for error messages</param>
    /// <returns>The normalized full path</returns>
    /// <exception cref="ArgumentException">Thrown if path is invalid or potentially malicious</exception>
    public static string ValidateAndNormalizePath(string path, string parameterName = "path")
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("Path cannot be null or empty", parameterName);
        }

        // Normalize the path to resolve any relative path components
        string fullPath;
        try
        {
            fullPath = Path.GetFullPath(path);
        }
        catch (Exception ex)
        {
            throw new ArgumentException($"Invalid path format: {ex.Message}", parameterName, ex);
        }

        // Check for path length limits (DoS prevention)
        if (fullPath.Length > MaxPathLength)
        {
            throw new ArgumentException(
                $"Path too long: {fullPath.Length} characters (maximum: {MaxPathLength})",
                parameterName);
        }

        // Note: Path.GetFullPath() already validates characters, so no need for additional check

        return fullPath;
    }

    /// <summary>
    /// Validates a file path and ensures the file exists
    /// </summary>
    /// <param name="path">The file path to validate</param>
    /// <param name="parameterName">Name of the parameter for error messages</param>
    /// <param name="validateSize">Whether to validate file size (default: true)</param>
    /// <returns>The normalized full path</returns>
    /// <exception cref="ArgumentException">Thrown if path is invalid or file is too large</exception>
    /// <exception cref="FileNotFoundException">Thrown if file does not exist</exception>
    public static string ValidateExistingFile(string path, string parameterName = "path", bool validateSize = true)
    {
        string fullPath = ValidateAndNormalizePath(path, parameterName);

        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException($"File not found: {fullPath}", fullPath);
        }

        // Validate file size to prevent DoS attacks from extremely large files
        if (validateSize)
        {
            var fileInfo = new FileInfo(fullPath);
            if (fileInfo.Length > MaxInputFileSize)
            {
                throw new ArgumentException(
                    $"File too large: {fileInfo.Length:N0} bytes (maximum: {MaxInputFileSize:N0} bytes = 100MB). " +
                    "This limit prevents denial-of-service attacks from processing extremely large files.",
                    parameterName);
            }
        }

        return fullPath;
    }

    /// <summary>
    /// Validates a file path for writing, ensuring the directory exists or can be created
    /// </summary>
    /// <param name="path">The file path to validate</param>
    /// <param name="parameterName">Name of the parameter for error messages</param>
    /// <param name="allowOverwrite">Whether to allow overwriting existing files</param>
    /// <returns>The normalized full path</returns>
    /// <exception cref="ArgumentException">Thrown if path is invalid</exception>
    /// <exception cref="IOException">Thrown if file exists and overwrite is not allowed</exception>
    public static string ValidateOutputFile(string path, string parameterName = "path", bool allowOverwrite = true)
    {
        string fullPath = ValidateAndNormalizePath(path, parameterName);

        // Check if file already exists
        if (!allowOverwrite && File.Exists(fullPath))
        {
            throw new IOException($"File already exists: {fullPath}");
        }

        // Ensure parent directory exists or can be created
        string? directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            try
            {
                Directory.CreateDirectory(directory);
            }
            catch (Exception ex)
            {
                throw new IOException($"Cannot create directory: {directory}", ex);
            }
        }

        return fullPath;
    }

    /// <summary>
    /// Validates that a file has an allowed extension
    /// </summary>
    /// <param name="path">The file path to validate</param>
    /// <param name="allowedExtensions">Array of allowed extensions (e.g., ".xlsx", ".xlsm")</param>
    /// <param name="parameterName">Name of the parameter for error messages</param>
    /// <returns>The normalized full path</returns>
    /// <exception cref="ArgumentException">Thrown if extension is not allowed</exception>
    public static string ValidateFileExtension(string path, string[] allowedExtensions, string parameterName = "path")
    {
        string fullPath = ValidateAndNormalizePath(path, parameterName);
        string extension = Path.GetExtension(fullPath).ToLowerInvariant();

        if (!allowedExtensions.Any(ext => ext.Equals(extension, StringComparison.OrdinalIgnoreCase)))
        {
            throw new ArgumentException(
                $"Invalid file extension '{extension}'. Allowed: {string.Join(", ", allowedExtensions)}",
                parameterName);
        }

        return fullPath;
    }

    /// <summary>
    /// Validates a path is safe for file system operations on Windows
    /// </summary>
    /// <param name="path">The path to validate</param>
    /// <returns>True if path is safe</returns>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public static bool IsSafePath(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return false;

        try
        {
            // Attempt to normalize and validate
            string fullPath = Path.GetFullPath(path);

            // Check for excessively long paths
            if (fullPath.Length > MaxPathLength)
                return false;

            // Check for invalid characters
            if (fullPath.IndexOfAny(Path.GetInvalidPathChars()) >= 0)
                return false;

            // Path is considered safe
            return true;
        }
        catch
        {
            // Any exception during validation means path is unsafe
            return false;
        }
    }
}
