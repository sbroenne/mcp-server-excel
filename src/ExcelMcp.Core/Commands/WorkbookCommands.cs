using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Session;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Implementation of workbook lifecycle management commands.
/// Provides explicit control over opening, saving, and closing Excel workbooks using the active workbook pattern.
/// </summary>
public class WorkbookCommands : IWorkbookCommands
{
    private readonly IFileCommands _fileCommands;

    /// <summary>
    /// Creates a new instance of WorkbookCommands
    /// </summary>
    /// <param name="fileCommands">File commands for creating new workbooks</param>
    public WorkbookCommands(IFileCommands fileCommands)
    {
        _fileCommands = fileCommands ?? throw new ArgumentNullException(nameof(fileCommands));
    }

    /// <inheritdoc/>
    public async Task<OperationResult> CreateAsync(string filePath, bool overwriteIfExists = false)
    {
        try
        {
            // First, create the empty workbook file using existing FileCommands
            var createResult = await _fileCommands.CreateEmptyAsync(filePath, overwriteIfExists);
            if (!createResult.Success)
            {
                return new OperationResult
                {
                    Success = false,
                    ErrorMessage = createResult.ErrorMessage,
                    FilePath = filePath
                };
            }

            // Now open it and set as active workbook
            var handle = await FileHandleManager.OpenAsync(filePath);

            return new OperationResult
            {
                Success = true,
                FilePath = filePath,
                Action = "create",
                OperationContext = new Dictionary<string, object>
                {
                    ["HandleId"] = handle.Id,
                    ["OpenedAt"] = handle.OpenedAt
                }
            };
        }
        catch (Exception ex)
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"Failed to create workbook: {ex.Message}",
                FilePath = filePath
            };
        }
    }

    /// <inheritdoc/>
    public async Task<OperationResult> OpenAsync(string filePath)
    {
        try
        {
            // Validate file exists first
            if (!File.Exists(filePath))
            {
                return new OperationResult
                {
                    Success = false,
                    ErrorMessage = $"File not found: {filePath}",
                    FilePath = filePath
                };
            }

            // Open and set as active workbook (reuses handle if already open)
            var handle = await FileHandleManager.OpenAsync(filePath);

            return new OperationResult
            {
                Success = true,
                FilePath = filePath,
                Action = "open",
                OperationContext = new Dictionary<string, object>
                {
                    ["HandleId"] = handle.Id,
                    ["OpenedAt"] = handle.OpenedAt,
                    ["IsReused"] = true  // Could track this in manager if needed
                }
            };
        }
        catch (Exception ex)
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"Failed to open workbook: {ex.Message}",
                FilePath = filePath
            };
        }
    }

    /// <inheritdoc/>
    public async Task<OperationResult> SaveAsync(string? filePath = null)
    {
        try
        {
            string targetPath = filePath ?? ActiveWorkbook.Current.FilePath;

            await FileHandleManager.SaveAsync(filePath);

            return new OperationResult
            {
                Success = true,
                FilePath = targetPath,
                Action = "save"
            };
        }
        catch (InvalidOperationException ex)
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                FilePath = filePath
            };
        }
        catch (Exception ex)
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"Failed to save workbook: {ex.Message}",
                FilePath = filePath
            };
        }
    }

    /// <inheritdoc/>
    public async Task<OperationResult> CloseAsync(string? filePath = null)
    {
        try
        {
            string targetPath = filePath ?? ActiveWorkbook.Current.FilePath;

            await FileHandleManager.CloseAsync(filePath);

            return new OperationResult
            {
                Success = true,
                FilePath = targetPath,
                Action = "close"
            };
        }
        catch (InvalidOperationException ex)
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                FilePath = filePath
            };
        }
        catch (Exception ex)
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"Failed to close workbook: {ex.Message}",
                FilePath = filePath
            };
        }
    }

    /// <inheritdoc/>
    public async Task<OperationResult> GetActiveWorkbookAsync()
    {
        return await Task.Run(() =>
        {
            try
            {
                if (!ActiveWorkbook.HasActive)
                {
                    return new OperationResult
                    {
                        Success = false,
                        ErrorMessage = "No active workbook. Call OpenAsync() or CreateAsync() first."
                    };
                }

                var handle = ActiveWorkbook.Current;

                return new OperationResult
                {
                    Success = true,
                    FilePath = handle.FilePath,
                    Action = "get-active",
                    OperationContext = new Dictionary<string, object>
                    {
                        ["HandleId"] = handle.Id,
                        ["OpenedAt"] = handle.OpenedAt,
                        ["IsClosed"] = handle.IsClosed
                    }
                };
            }
            catch (Exception ex)
            {
                return new OperationResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to get active workbook: {ex.Message}"
                };
            }
        });
    }
}
