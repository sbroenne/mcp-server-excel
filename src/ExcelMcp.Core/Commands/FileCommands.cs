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
        filePath = Path.GetFullPath(filePath);

        bool exists = File.Exists(filePath);
        string extension = Path.GetExtension(filePath).ToLowerInvariant();
        bool isValidExtension = extension is ".xlsx" or ".xlsm";

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

        return new FileValidationInfo
        {
            FilePath = filePath,
            Exists = exists,
            Size = size,
            Extension = extension,
            LastModified = lastModified,
            IsValid = exists && isValidExtension,
            Message = message
        };
    }

}




