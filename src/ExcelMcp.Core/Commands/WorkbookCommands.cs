using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Implementation of workbook lifecycle commands using FileHandleManager.
/// </summary>
public sealed class WorkbookCommands : IWorkbookCommands
{
    /// <summary>
    /// Creates a new Excel workbook file and caches its handle.
    /// </summary>
    public async Task<OperationResult> CreateAsync(string filePath)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return new OperationResult
                {
                    Success = false,
                    ErrorMessage = "File path is required"
                };
            }

            var handle = await FileHandleManager.Instance.CreateAsync(filePath);

            return new OperationResult
            {
                Success = true,
                FilePath = handle.FilePath,
                Action = "create"
            };
        }
        catch (Exception ex)
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"Failed to create workbook: {ex.Message}"
            };
        }
    }

    /// <summary>
    /// Opens an existing workbook file or retrieves cached handle if already open.
    /// </summary>
    public async Task<OperationResult> OpenAsync(string filePath)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return new OperationResult
                {
                    Success = false,
                    ErrorMessage = "File path is required"
                };
            }

            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            return new OperationResult
            {
                Success = true,
                FilePath = handle.FilePath,
                Action = "open"
            };
        }
        catch (Exception ex)
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"Failed to open workbook: {ex.Message}"
            };
        }
    }

    /// <summary>
    /// Saves the workbook to disk.
    /// </summary>
    public async Task<OperationResult> SaveAsync(string filePath)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return new OperationResult
                {
                    Success = false,
                    ErrorMessage = "File path is required"
                };
            }

            await FileHandleManager.Instance.SaveAsync(filePath);

            return new OperationResult
            {
                Success = true,
                FilePath = Path.GetFullPath(filePath),
                Action = "save"
            };
        }
        catch (InvalidOperationException ex)
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"File not open: {ex.Message}"
            };
        }
        catch (Exception ex)
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"Failed to save workbook: {ex.Message}"
            };
        }
    }

    /// <summary>
    /// Explicitly closes the workbook and releases its handle from cache.
    /// </summary>
    public async Task<OperationResult> CloseAsync(string filePath)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return new OperationResult
                {
                    Success = false,
                    ErrorMessage = "File path is required"
                };
            }

            await FileHandleManager.Instance.CloseAsync(filePath);

            return new OperationResult
            {
                Success = true,
                FilePath = Path.GetFullPath(filePath),
                Action = "close"
            };
        }
        catch (Exception ex)
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"Failed to close workbook: {ex.Message}"
            };
        }
    }

    /// <summary>
    /// Gets a list of all currently open files (cached handles).
    /// </summary>
    public Task<OperationResult> ListOpenFilesAsync()
    {
        try
        {
            var openFiles = FileHandleManager.Instance.GetOpenFiles();

            return Task.FromResult(new OperationResult
            {
                Success = true,
                Action = "list-open-files"
                // Note: File list is accessible via FileHandleManager.Instance.GetOpenFiles()
                // For now, we just return success. A future iteration could create
                // a WorkbookListResult type if needed.
            });
        }
        catch (Exception ex)
        {
            return Task.FromResult(new OperationResult
            {
                Success = false,
                ErrorMessage = $"Failed to list open files: {ex.Message}",
                Action = "list-open-files"
            });
        }
    }
}
