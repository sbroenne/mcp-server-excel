using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;


namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// File management commands implementation
/// </summary>
public class FileCommands : IFileCommands
{
    /// <inheritdoc />
    public OperationResult CreateEmpty(string filePath, bool overwriteIfExists = false)
    {
        try
        {
            filePath = Path.GetFullPath(filePath);

            // Validate file extension
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension is not ".xlsx" and not ".xlsm")
            {
                return new OperationResult
                {
                    Success = false,
                    ErrorMessage = "File must have .xlsx or .xlsm extension",
                    FilePath = filePath,
                    Action = "create-empty"
                };
            }

            // Check if file already exists
            if (File.Exists(filePath) && !overwriteIfExists)
            {
                return new OperationResult
                {
                    Success = false,
                    ErrorMessage = $"File already exists: {filePath}",
                    FilePath = filePath,
                    Action = "create-empty"
                };
            }

            // Ensure directory exists
            string? directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                try
                {
                    Directory.CreateDirectory(directory);
                }
                catch (Exception ex)
                {
                    return new OperationResult
                    {
                        Success = false,
                        ErrorMessage = $"Failed to create directory: {ex.Message}",
                        FilePath = filePath,
                        Action = "create-empty"
                    };
                }
            }

            // Create Excel workbook using proper resource management
            bool isMacroEnabled = extension == ".xlsm";

            // CRITICAL: CreateNew opens a batch but does NOT save automatically
            // We must save the batch after making changes, then dispose
            var result = ExcelSession.CreateNew(filePath, isMacroEnabled, (ctx, ct) =>
            {
                // Set up a basic structure with proper COM cleanup
                dynamic? sheet = null;
                dynamic? cell = null;
                dynamic? comment = null;

                try
                {
                    sheet = ctx.Book.Worksheets.Item(1);
                    sheet.Name = "Sheet1";

                    // Add a comment to indicate this was created by ExcelCLI
                    cell = sheet.Range["A1"];
                    comment = cell.AddComment($"Created by ExcelCLI on {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    comment.Visible = false;

                    // Save changes before the batch is disposed
                    ctx.Book.Save();

                    return new OperationResult
                    {
                        Success = true,
                        FilePath = filePath,
                        Action = "create-empty"
                    };
                }
                finally
                {
                    ComUtilities.Release(ref comment);
                    ComUtilities.Release(ref cell);
                    ComUtilities.Release(ref sheet);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"Failed to create Excel file: {ex.Message}",
                FilePath = filePath,
                Action = "create-empty"
            };
        }
    }

    /// <inheritdoc />
    public FileValidationResult Test(string filePath)
    {
        try
        {
            filePath = Path.GetFullPath(filePath);

            // Check if file exists
            bool exists = File.Exists(filePath);

            // Get file extension
            string extension = exists ? Path.GetExtension(filePath).ToLowerInvariant() : "";

            // Validate extension
            bool isValidExtension = extension is ".xlsx" or ".xlsm";

            // Get file info if exists
            long size = 0;
            DateTime lastModified = DateTime.MinValue;

            if (exists)
            {
                var fileInfo = new FileInfo(filePath);
                size = fileInfo.Length;
                lastModified = fileInfo.LastWriteTime;
            }

            return new FileValidationResult
            {
                Success = exists && isValidExtension,
                ErrorMessage = !exists ? $"File not found: {filePath}"
                    : !isValidExtension ? $"Invalid file extension. Expected .xlsx or .xlsm, got {extension}"
                    : null,
                FilePath = filePath,
                Exists = exists,
                Size = size,
                Extension = extension,
                LastModified = lastModified,
                IsValid = exists && isValidExtension
            };
        }
        catch (Exception ex)
        {
            return new FileValidationResult
            {
                Success = false,
                ErrorMessage = $"Failed to validate file: {ex.Message}",
                FilePath = filePath,
                Exists = false,
                Size = 0,
                Extension = "",
                LastModified = DateTime.MinValue,
                IsValid = false
            };
        }
    }

}


