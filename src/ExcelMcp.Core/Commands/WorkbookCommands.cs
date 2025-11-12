using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Implementation of workbook lifecycle commands using FileHandleManager.
/// </summary>
public class WorkbookCommands : IWorkbookCommands
{
    /// <inheritdoc />
    public async Task<OperationResult> CreateAsync(string filePath)
    {
        var result = new OperationResult { Action = "create-workbook" };

        try
        {
            // Validate path
            if (string.IsNullOrWhiteSpace(filePath))
            {
                result.Success = false;
                result.ErrorMessage = "File path cannot be empty";
                return result;
            }

            string fullPath = Path.GetFullPath(filePath);
            result.FilePath = fullPath;

            // Check if file already exists
            if (File.Exists(fullPath))
            {
                result.Success = false;
                result.ErrorMessage = $"File already exists: {fullPath}";
                return result;
            }

            // Create directory if needed
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Open or get handle (will create new workbook if doesn't exist)
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(fullPath);
            await handle.SaveAsync(); // Ensure it's saved to disk

            result.Success = true;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to create workbook: {ex.Message}";
        }

        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> OpenAsync(string filePath)
    {
        var result = new OperationResult { Action = "open-workbook" };

        try
        {
            // Validate path
            if (string.IsNullOrWhiteSpace(filePath))
            {
                result.Success = false;
                result.ErrorMessage = "File path cannot be empty";
                return result;
            }

            string fullPath = Path.GetFullPath(filePath);
            result.FilePath = fullPath;

            // Check if file exists
            if (!File.Exists(fullPath))
            {
                result.Success = false;
                result.ErrorMessage = $"File not found: {fullPath}";
                return result;
            }

            // Open or get cached handle
            await FileHandleManager.Instance.OpenOrGetAsync(fullPath);

            result.Success = true;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to open workbook: {ex.Message}";
        }

        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> SaveAsync(string filePath)
    {
        var result = new OperationResult { Action = "save-workbook" };

        try
        {
            // Validate path
            if (string.IsNullOrWhiteSpace(filePath))
            {
                result.Success = false;
                result.ErrorMessage = "File path cannot be empty";
                return result;
            }

            string fullPath = Path.GetFullPath(filePath);
            result.FilePath = fullPath;

            // Check if handle exists
            if (!FileHandleManager.Instance.HasHandle(fullPath))
            {
                result.Success = false;
                result.ErrorMessage = $"Workbook not open: {fullPath}. Call OpenAsync or another operation first.";
                return result;
            }

            // Save via FileHandleManager
            await FileHandleManager.Instance.SaveAsync(fullPath);

            result.Success = true;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to save workbook: {ex.Message}";
        }

        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> CloseAsync(string filePath)
    {
        var result = new OperationResult { Action = "close-workbook" };

        try
        {
            // Validate path
            if (string.IsNullOrWhiteSpace(filePath))
            {
                result.Success = false;
                result.ErrorMessage = "File path cannot be empty";
                return result;
            }

            string fullPath = Path.GetFullPath(filePath);
            result.FilePath = fullPath;

            // Close handle (no error if not open)
            await FileHandleManager.Instance.CloseAsync(fullPath);

            result.Success = true;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to close workbook: {ex.Message}";
        }

        return result;
    }

    /// <inheritdoc />
    public Task<OperationResult> ListOpenFilesAsync()
    {
        var result = new OperationResult { Action = "list-open-files" };

        try
        {
            var openFiles = FileHandleManager.Instance.GetOpenFiles();

            result.Success = true;
            // Store open files count in OperationContext
            result.OperationContext = new Dictionary<string, object>
            {
                ["OpenFileCount"] = openFiles.Count,
                ["OpenFiles"] = openFiles
            };
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to list open files: {ex.Message}";
        }

        return Task.FromResult(result);
    }
}
