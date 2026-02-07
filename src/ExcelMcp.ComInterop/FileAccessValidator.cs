namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>
/// Utility class for validating file access and locking status.
/// Provides OS-level file lock detection before Excel COM operations.
/// </summary>
public static class FileAccessValidator
{
    /// <summary>
    /// Validates that a file is not locked by attempting to open it with exclusive access.
    /// Throws InvalidOperationException if file is locked or inaccessible.
    /// This is a fast OS-level check that doesn't require launching Excel.
    /// </summary>
    /// <param name="filePath">The file path to validate</param>
    /// <exception cref="InvalidOperationException">Thrown when file is locked or inaccessible</exception>
    public static void ValidateFileNotLocked(string filePath)
    {
        try
        {
            using var lockTest = new FileStream(
                filePath,
                FileMode.Open,
                FileAccess.ReadWrite,
                FileShare.None);
            // File is NOT locked - close and proceed
        }
        catch (IOException ioEx)
        {
            // File is locked by another process (most likely already open in Excel)
            throw CreateFileLockedError(filePath, ioEx);
        }
        catch (UnauthorizedAccessException uaEx)
        {
            // File access denied (permissions issue or file is locked)
            throw new InvalidOperationException(
                $"Cannot access '{Path.GetFileName(filePath)}'. " +
                "The file may be read-only, you may lack permissions, or it's locked by another process. " +
                "Please verify file permissions and close any applications using this file.",
                uaEx);
        }
    }

    /// <summary>
    /// Creates a standardized InvalidOperationException for file-locked scenarios.
    /// Provides consistent error messages across the codebase.
    /// </summary>
    /// <param name="filePath">The file path that is locked</param>
    /// <param name="innerException">The underlying exception that triggered the error</param>
    /// <returns>A user-friendly InvalidOperationException with guidance</returns>
    public static InvalidOperationException CreateFileLockedError(string filePath, Exception innerException)
    {
        return new InvalidOperationException(
            $"Cannot open '{Path.GetFileName(filePath)}'. " +
            "The file is already open in Excel or another process is using it. " +
            "Please close the file before running automation commands. " +
            "ExcelMcp requires exclusive access to workbooks during operations.",
            innerException);
    }
}


