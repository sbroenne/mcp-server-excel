using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// File management commands implementation
/// </summary>
public class FileCommands : IFileCommands
{
    /// <inheritdoc />
    public FileValidationInfo Test(string filePath)
    {
        filePath = FilePathValidation.NormalizeAbsolutePath(filePath);

        bool exists = File.Exists(filePath);
        string extension = Path.GetExtension(filePath).ToLowerInvariant();
        bool isValidExtension = extension is ".xlsx" or ".xlsm";
        bool isIrmProtected = exists && isValidExtension && FileAccessValidator.IsIrmProtected(filePath);
        bool isValid = false;
        bool canOpen = false;
        bool preflightPassed = false;

        long size = 0;
        DateTime lastModified = DateTime.MinValue;

        if (exists)
        {
            var fileInfo = new FileInfo(filePath);
            size = fileInfo.Length;
            lastModified = fileInfo.LastWriteTime;
        }

        string? message = !exists
            ? $"File not found: {filePath}"
            : !isValidExtension ? $"Invalid file extension. Expected .xlsx or .xlsm, got {extension}" : null;

        if (exists && isValidExtension)
        {
            try
            {
                if (isIrmProtected)
                {
                    using var readTest = new FileStream(
                        filePath,
                        FileMode.Open,
                        FileAccess.Read,
                        FileShare.ReadWrite);
                    message =
                        "IRM/AIP protection detected. Structural validity and openability require " +
                        "an interactive Excel open; use show=true. ExcelMcp will open this file read-only.";
                    preflightPassed = true;
                }
                else
                {
                    FileAccessValidator.ValidateFileNotLocked(filePath);
                    using var readTest = new FileStream(
                        filePath,
                        FileMode.Open,
                        FileAccess.Read,
                        FileShare.ReadWrite);
                    message =
                        "Path, extension, lock, and read-access checks passed. " +
                        "ExcelMcp treats workbook contents as opaque, so structural validity " +
                        "and openability were not inspected. Use file(action: 'open') or " +
                        "session open to have Excel validate the workbook.";
                    preflightPassed = true;
                }
            }
            catch (InvalidOperationException ex)
            {
                message = ex.Message;
            }
            catch (IOException ex)
            {
                message = $"Cannot read '{Path.GetFileName(filePath)}': {ex.Message}";
            }
            catch (UnauthorizedAccessException ex)
            {
                message = $"Cannot read '{Path.GetFileName(filePath)}': {ex.Message}";
            }
        }

        return new FileValidationInfo
        {
            FilePath = filePath,
            Exists = exists,
            Size = size,
            Extension = extension,
            LastModified = lastModified,
            PreflightPassed = preflightPassed,
            IsValid = isValid,
            CanOpen = canOpen,
            IsIrmProtected = isIrmProtected,
            WillOpenReadOnly = isIrmProtected,
            RequiresVisibleSession = isIrmProtected,
            Message = message
        };
    }

}
