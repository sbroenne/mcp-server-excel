using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// File management commands implementation
/// </summary>
public class FileCommands : IFileCommands
{
    /// <inheritdoc />
    public FileValidationInfo Test(string filePath)
    {
        var validation = ExcelFileValidator.Inspect(filePath);
        var lastModified = validation.Exists
            ? new FileInfo(validation.FilePath).LastWriteTime
            : DateTime.MinValue;

        return new FileValidationInfo
        {
            FilePath = validation.FilePath,
            Exists = validation.Exists,
            Size = validation.Size,
            Extension = validation.Extension,
            LastModified = lastModified,
            IsValid = validation.IsValidExistingWorkbook,
            IsIrmProtected = validation.Exists && FileAccessValidator.IsIrmProtected(validation.FilePath),
            Message = validation.Message
        };
    }

}




