using System.Collections.Concurrent;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Session;

/// <summary>
/// Internal file handle manager for the active workbook pattern.
/// Caches Excel instances by file path to enable handle reuse and eliminate file lock issues.
/// </summary>
internal sealed class FileHandleManager
{
    // Key by absolute file path - reuses handles for same file
    private static readonly ConcurrentDictionary<string, (FileHandle Handle, IExcelBatch Batch)> _handlesByPath = new();

    /// <summary>
    /// Opens an existing workbook file and sets it as the active workbook.
    /// If file is already open, reuses the existing handle and sets it as active.
    /// </summary>
    /// <param name="filePath">Path to the workbook file to open</param>
    /// <returns>File handle for the opened workbook (may be existing handle)</returns>
    internal static async Task<FileHandle> OpenAsync(string filePath)
    {
        string absolutePath = Path.GetFullPath(filePath);

        // Check if already open - reuse existing handle
        if (_handlesByPath.TryGetValue(absolutePath, out var existing))
        {
            ActiveWorkbook.Current = existing.Handle;
            return existing.Handle;  // Reuse existing handle (solves Issue #173!)
        }

        // Open new instance
        var handle = new FileHandle(Guid.NewGuid().ToString(), absolutePath);
        var batch = await ExcelSession.BeginBatchAsync(absolutePath);

        _handlesByPath[absolutePath] = (handle, batch);
        ActiveWorkbook.Current = handle;

        return handle;
    }

    /// <summary>
    /// Gets the Excel batch for the active workbook.
    /// </summary>
    /// <returns>Excel batch instance for executing operations</returns>
    internal static IExcelBatch GetActiveWorkbookBatch()
    {
        var handle = ActiveWorkbook.Current;

        // Find by file path
        if (!_handlesByPath.TryGetValue(handle.FilePath, out var entry))
        {
            throw new InvalidOperationException($"Active workbook not found: {handle.FilePath}");
        }

        return entry.Batch;
    }

    /// <summary>
    /// Saves changes to the active workbook or a specific workbook.
    /// </summary>
    /// <param name="filePath">Optional file path. If null, saves the active workbook.</param>
    internal static async Task SaveAsync(string? filePath = null)
    {
        string targetPath = filePath != null
            ? Path.GetFullPath(filePath)
            : ActiveWorkbook.Current.FilePath;

        if (!_handlesByPath.TryGetValue(targetPath, out var entry))
        {
            throw new InvalidOperationException($"Workbook not found: {targetPath}");
        }

        await entry.Batch.SaveAsync();
    }

    /// <summary>
    /// Closes a workbook and releases its resources.
    /// If filePath is specified, closes that file. Otherwise, closes the active workbook.
    /// </summary>
    /// <param name="filePath">Optional file path. If null, closes the active workbook.</param>
    internal static async Task CloseAsync(string? filePath = null)
    {
        // Determine target path
        string targetPath = filePath != null
            ? Path.GetFullPath(filePath)
            : ActiveWorkbook.Current.FilePath;

        // Remove from cache
        if (!_handlesByPath.TryRemove(targetPath, out var entry))
        {
            return;  // Already closed or never opened
        }

        try
        {
            // Dispose the batch (closes workbook, quits Excel, cleanup COM)
            await entry.Batch.DisposeAsync();
        }
        finally
        {
            entry.Handle.IsClosed = true;

            // Clear active workbook if this was it
            if (ActiveWorkbook.HasActive && ActiveWorkbook.Current.FilePath == targetPath)
            {
                // Can't set to null directly due to AsyncLocal semantics
                // The caller should not access ActiveWorkbook after closing it
            }
        }
    }

    /// <summary>
    /// Closes all open workbooks. Used for cleanup and testing.
    /// </summary>
    internal static async Task CloseAllAsync()
    {
        var allPaths = _handlesByPath.Keys.ToList();
        foreach (var path in allPaths)
        {
            await CloseAsync(path);
        }
    }
}
