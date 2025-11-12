using System.Collections.Concurrent;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Singleton that caches Excel workbook handles by absolute file path.
/// Enables automatic handle reuse across multiple operations on the same file.
/// Thread-safe for concurrent operations.
/// </summary>
public sealed class FileHandleManager : IDisposable
{
    // Singleton instance - one per process (CLI or MCP Server)
    private static readonly Lazy<FileHandleManager> _instance =
        new(() => new FileHandleManager());

    /// <summary>
    /// Gets the singleton instance of the FileHandleManager.
    /// </summary>
    public static FileHandleManager Instance => _instance.Value;

    // Cache handles by absolute file path
    private readonly ConcurrentDictionary<string, ExcelWorkbookHandle> _handles = new();
    private readonly SemaphoreSlim _lock = new(1, 1);
    private bool _disposed;

    private FileHandleManager()
    {
        // Private constructor for singleton
    }

    /// <summary>
    /// Opens or retrieves cached handle for file path. Thread-safe.
    /// If the file is already open, returns the existing handle and resets its inactivity timeout.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the workbook file</param>
    /// <returns>Handle to the workbook (existing or newly opened)</returns>
    public async Task<ExcelWorkbookHandle> OpenOrGetAsync(string filePath)
    {
        ObjectDisposedException.ThrowIf(_disposed, nameof(FileHandleManager));

        string absolutePath = Path.GetFullPath(filePath);

        await _lock.WaitAsync();
        try
        {
            // Reuse existing handle if already open
            if (_handles.TryGetValue(absolutePath, out var existing))
            {
                existing.UpdateLastAccess();  // Reset inactivity timeout
                return existing;
            }

            // Create new handle by opening the workbook
            var newHandle = await ExcelWorkbookHandle.OpenAsync(absolutePath);
            _handles[absolutePath] = newHandle;
            return newHandle;
        }
        finally
        {
            _lock.Release();
        }
    }

    /// <summary>
    /// Creates a new workbook and caches its handle.
    /// </summary>
    /// <param name="filePath">Path where the new workbook should be created</param>
    /// <returns>Handle to the newly created workbook</returns>
    /// <exception cref="InvalidOperationException">If file already exists or is already cached</exception>
    public async Task<ExcelWorkbookHandle> CreateAsync(string filePath)
    {
        ObjectDisposedException.ThrowIf(_disposed, nameof(FileHandleManager));

        string absolutePath = Path.GetFullPath(filePath);

        await _lock.WaitAsync();
        try
        {
            // Check if already cached
            if (_handles.ContainsKey(absolutePath))
            {
                throw new InvalidOperationException($"File is already open: {absolutePath}");
            }

            // Create new workbook
            var newHandle = await ExcelWorkbookHandle.CreateAsync(absolutePath);
            _handles[absolutePath] = newHandle;
            return newHandle;
        }
        finally
        {
            _lock.Release();
        }
    }

    /// <summary>
    /// Gets handle for file path (must already be open).
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the workbook file</param>
    /// <returns>The cached handle</returns>
    /// <exception cref="InvalidOperationException">If file is not currently open</exception>
    public ExcelWorkbookHandle GetHandle(string filePath)
    {
        ObjectDisposedException.ThrowIf(_disposed, nameof(FileHandleManager));

        string absolutePath = Path.GetFullPath(filePath);

        if (_handles.TryGetValue(absolutePath, out var handle))
        {
            handle.UpdateLastAccess();
            return handle;
        }

        throw new InvalidOperationException($"File not open: {filePath}. Call OpenOrGetAsync() or CreateAsync() first.");
    }

    /// <summary>
    /// Explicitly closes handle for file path. Removes from cache and disposes.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the workbook file</param>
    public async Task CloseAsync(string filePath)
    {
        if (_disposed)
        {
            return; // Already disposed, nothing to do
        }

        string absolutePath = Path.GetFullPath(filePath);

        await _lock.WaitAsync();
        try
        {
            if (_handles.TryRemove(absolutePath, out var handle))
            {
                await handle.DisposeAsync();
            }
        }
        finally
        {
            _lock.Release();
        }
    }

    /// <summary>
    /// Saves specific file (must already be open).
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the workbook file</param>
    /// <exception cref="InvalidOperationException">If file is not currently open</exception>
    public async Task SaveAsync(string filePath)
    {
        var handle = GetHandle(filePath);  // Throws if not open
        await handle.SaveAsync();
    }

    /// <summary>
    /// Gets a list of all currently cached file paths.
    /// Useful for diagnostics and debugging.
    /// </summary>
    /// <returns>Array of absolute file paths for all cached handles</returns>
    public string[] GetOpenFiles()
    {
        if (_disposed)
        {
            return Array.Empty<string>();
        }

        return _handles.Keys.ToArray();
    }

    /// <summary>
    /// Background cleanup: Close handles inactive for longer than the specified timeout.
    /// Runs periodically (e.g., every minute) to prevent resource leaks.
    /// </summary>
    /// <param name="inactivityTimeout">Time span after which inactive handles are closed</param>
    /// <returns>Number of handles closed</returns>
    public async Task<int> CleanupInactiveHandlesAsync(TimeSpan inactivityTimeout)
    {
        if (_disposed)
        {
            return 0;
        }

        var now = DateTime.UtcNow;
        var toRemove = _handles
            .Where(kvp => (now - kvp.Value.LastAccess) > inactivityTimeout)
            .Select(kvp => kvp.Key)
            .ToList();

        foreach (var path in toRemove)
        {
            await CloseAsync(path);
        }

        return toRemove.Count;
    }

    /// <summary>
    /// Closes all cached handles and clears the cache.
    /// </summary>
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        // Dispose all handles
        foreach (var handle in _handles.Values)
        {
            try
            {
                handle.Dispose();
            }
            catch
            {
                // Suppress disposal errors
            }
        }

        _handles.Clear();
        _lock.Dispose();
    }
}
