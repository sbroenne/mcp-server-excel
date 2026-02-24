namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>
/// Utility class for validating file access and locking status.
/// Provides OS-level file lock detection and IRM/AIP-encryption detection before Excel COM operations.
/// </summary>
public static class FileAccessValidator
{
    // OLE2 Compound Document Format signature.
    // IRM/AIP-protected Excel files are stored as OLE2 containers with an EncryptedPackage
    // stream instead of the standard ZIP-based Office Open XML format.
    private static ReadOnlySpan<byte> Ole2Signature => [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

    /// <summary>
    /// Detects if the file is IRM/AIP-protected by checking for the OLE2 compound document
    /// signature. IRM-protected files must be opened as read-only with Excel visible so the
    /// user can authenticate through the Information Rights Management credential prompt.
    /// </summary>
    /// <param name="filePath">The file path to inspect.</param>
    /// <returns>
    /// <c>true</c> if the file has the OLE2 Compound Document header, indicating IRM/AIP
    /// encryption; <c>false</c> for standard ZIP-based .xlsx/.xlsm files or if the file
    /// cannot be read.
    /// </returns>
    public static bool IsIrmProtected(string filePath)
    {
        if (!File.Exists(filePath))
            return false;
        try
        {
            Span<byte> header = stackalloc byte[8];
            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            int read = fs.Read(header);
            if (read < 8)
                return false;
            return header.SequenceEqual(Ole2Signature);
        }
        catch
        {
            // Cannot read → treat as not IRM so normal error handling takes over
            return false;
        }
    }

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


