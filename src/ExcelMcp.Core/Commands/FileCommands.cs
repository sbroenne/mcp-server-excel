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
        filePath = FilePathValidation.NormalizeWorkbookLocation(filePath);
        if (WorkbookLocation.IsRemote(filePath))
        {
            return new FileValidationInfo
            {
                FilePath = filePath,
                Extension = WorkbookLocation.GetExtension(filePath),
                Message = "Remote file existence, size and IRM/AIP protection cannot be checked locally. " +
                    "Open in Excel to check access. Use show=true if you need to answer authentication prompts. " +
                    "False exists/isIrmProtected values do not prove that the file is missing or unprotected."
            };
        }

        bool exists = File.Exists(filePath);
        string extension = Path.GetExtension(filePath).ToLowerInvariant();
        bool isIrmProtected = exists && FileAccessValidator.IsIrmProtected(filePath);
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
            : null;

        if (exists && isIrmProtected)
        {
            try
            {
                using var readTest = new FileStream(
                    filePath,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.ReadWrite);
                message =
                    "IRM/AIP protection detected. Container validity and openability require " +
                    "an interactive Excel open; use show=true. Excel determines editing permissions " +
                    "after authentication; inspect readOnly using workbook get-info.";
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
            WillOpenReadOnly = false,
            RequiresVisibleSession = isIrmProtected,
            Message = message
        };
    }

}
