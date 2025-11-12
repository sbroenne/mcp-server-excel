using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Commands for workbook lifecycle management (create, open, save, close).
/// Provides explicit control over workbook handles and file operations.
/// </summary>
public interface IWorkbookCommands
{
    /// <summary>
    /// Creates a new Excel workbook at the specified path.
    /// Opens a handle in FileHandleManager for subsequent operations.
    /// </summary>
    /// <param name="filePath">Path where the new workbook will be created</param>
    /// <returns>Operation result indicating success or failure</returns>
    Task<OperationResult> CreateAsync(string filePath);

    /// <summary>
    /// Explicitly opens an existing workbook (optional - FileHandleManager auto-opens on demand).
    /// Useful for pre-loading a workbook handle.
    /// </summary>
    /// <param name="filePath">Path to the workbook to open</param>
    /// <returns>Operation result indicating success or failure</returns>
    Task<OperationResult> OpenAsync(string filePath);

    /// <summary>
    /// Saves the workbook (must already be open via FileHandleManager).
    /// </summary>
    /// <param name="filePath">Path to the workbook to save</param>
    /// <returns>Operation result indicating success or failure</returns>
    Task<OperationResult> SaveAsync(string filePath);

    /// <summary>
    /// Explicitly closes the workbook and releases the handle from FileHandleManager.
    /// Optional - handles are auto-closed after inactivity timeout (5 minutes).
    /// </summary>
    /// <param name="filePath">Path to the workbook to close</param>
    /// <returns>Operation result indicating success or failure</returns>
    Task<OperationResult> CloseAsync(string filePath);

    /// <summary>
    /// Gets list of currently open workbook file paths.
    /// Useful for diagnostics and troubleshooting.
    /// </summary>
    /// <returns>Operation result with list of open file paths</returns>
    Task<OperationResult> ListOpenFilesAsync();
}
