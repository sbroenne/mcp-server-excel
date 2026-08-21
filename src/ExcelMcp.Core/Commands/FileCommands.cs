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
        filePath = FilePathValidation.NormalizeAbsoluteWindowsPath(filePath);

        bool exists = File.Exists(filePath);
        string extension = Path.GetExtension(filePath).ToLowerInvariant();
        bool isValidExtension = extension is ".xlsx" or ".xlsm";
        bool isIrmProtected = exists && isValidExtension && FileAccessValidator.IsIrmProtected(filePath);
        bool isValid = false;
        bool canOpen = false;

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
                        "IRM/AIP protection detected. Container validity and openability require " +
                        "an interactive Excel open; use show=true. ExcelMcp will open this file read-only.";
                }
                else
                {
                    FileAccessValidator.ValidateFileNotLocked(filePath);
                    isValid = FileAccessValidator.HasValidWorkbookContainer(filePath);
                    canOpen = isValid;
                    if (!isValid)
                    {
                        message = $"File is not a valid Excel workbook container: {filePath}";
                    }
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
            IsValid = isValid,
            CanOpen = canOpen,
            IsIrmProtected = isIrmProtected,
            WillOpenReadOnly = isIrmProtected,
            RequiresVisibleSession = isIrmProtected,
            Message = message
        };
    }

}
