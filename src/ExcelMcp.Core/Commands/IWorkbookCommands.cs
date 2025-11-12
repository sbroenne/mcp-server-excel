using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Workbook lifecycle commands for explicit file handle management.
/// NEW interface for filePath-based API (replaces batch-based workflow).
/// </summary>
public interface IWorkbookCommands
{
    /// <summary>
    /// Creates a new Excel workbook file and caches its handle.
    /// </summary>
    /// <param name="filePath">Path where the new workbook should be created</param>
    /// <returns>Operation result with handle information</returns>
    Task<OperationResult> CreateAsync(string filePath);

    /// <summary>
    /// Opens an existing workbook file or retrieves cached handle if already open.
    /// This is optional - most operations call OpenOrGetAsync internally.
    /// </summary>
    /// <param name="filePath">Path to the workbook file</param>
    /// <returns>Operation result with handle information</returns>
    Task<OperationResult> OpenAsync(string filePath);

    /// <summary>
    /// Saves the workbook to disk.
    /// File must already be open (via CreateAsync, OpenAsync, or any other operation).
    /// </summary>
    /// <param name="filePath">Path to the workbook file</param>
    /// <returns>Operation result</returns>
    Task<OperationResult> SaveAsync(string filePath);

    /// <summary>
    /// Explicitly closes the workbook and releases its handle from cache.
    /// After closing, the file is no longer accessible until reopened.
    /// </summary>
    /// <param name="filePath">Path to the workbook file</param>
    /// <returns>Operation result</returns>
    Task<OperationResult> CloseAsync(string filePath);

    /// <summary>
    /// Gets a list of all currently open files (cached handles).
    /// Useful for diagnostics and debugging.
    /// </summary>
    /// <returns>Result with list of open file paths</returns>
    Task<OperationResult> ListOpenFilesAsync();
}
