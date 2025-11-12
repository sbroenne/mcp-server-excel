using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Workbook lifecycle management commands for the active workbook pattern.
/// Provides explicit control over opening, saving, and closing Excel workbooks.
/// </summary>
public interface IWorkbookCommands
{
    /// <summary>
    /// Creates a new empty Excel workbook and sets it as the active workbook.
    /// </summary>
    /// <param name="filePath">Path where the workbook should be created</param>
    /// <param name="overwriteIfExists">Whether to overwrite if file already exists</param>
    /// <returns>Operation result with file handle information</returns>
    Task<OperationResult> CreateAsync(string filePath, bool overwriteIfExists = false);

    /// <summary>
    /// Opens an existing Excel workbook and sets it as the active workbook.
    /// If the file is already open, reuses the existing handle and sets it as active.
    /// This automatic handle reuse eliminates file lock issues from sequential operations.
    /// </summary>
    /// <param name="filePath">Path to the workbook file to open</param>
    /// <returns>Operation result with file handle information</returns>
    Task<OperationResult> OpenAsync(string filePath);

    /// <summary>
    /// Saves changes to the active workbook or a specific workbook.
    /// </summary>
    /// <param name="filePath">Optional file path. If null, saves the active workbook.</param>
    /// <returns>Operation result</returns>
    Task<OperationResult> SaveAsync(string? filePath = null);

    /// <summary>
    /// Closes the active workbook or a specific workbook and releases its resources.
    /// </summary>
    /// <param name="filePath">Optional file path. If null, closes the active workbook.</param>
    /// <returns>Operation result</returns>
    Task<OperationResult> CloseAsync(string? filePath = null);

    /// <summary>
    /// Gets information about the active workbook.
    /// </summary>
    /// <returns>Operation result with active workbook information</returns>
    Task<OperationResult> GetActiveWorkbookAsync();
}
