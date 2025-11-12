using System.Collections.Concurrent;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Manages Excel workbook handles with automatic caching and reuse by file path.
/// Thread-safe singleton for handle lifecycle management.
/// </summary>
public sealed class FileHandleManager : IDisposable
{
    // Singleton instance - one per process (CLI or MCP Server)
    private static readonly Lazy<FileHandleManager> _instance = new(() => new FileHandleManager());

    /// <summary>
    /// Singleton instance of FileHandleManager
    /// </summary>
    public static FileHandleManager Instance => _instance.Value;

    // Cache handles by absolute file path
    private readonly ConcurrentDictionary<string, ExcelWorkbookHandle> _handles = new();
    private readonly SemaphoreSlim _lock = new(1, 1);
    private bool _disposed;

    // Background cleanup configuration
    private readonly Timer? _cleanupTimer;
    private static readonly TimeSpan CleanupInterval = TimeSpan.FromMinutes(1);
    private static readonly TimeSpan InactivityTimeout = TimeSpan.FromMinutes(5);

    private FileHandleManager()
    {
        // Start background cleanup task
        _cleanupTimer = new Timer(
            async _ => await CleanupInactiveHandlesAsync(),
            null,
            CleanupInterval,
            CleanupInterval);
    }

    /// <summary>
    /// Opens or retrieves cached handle for file path. Thread-safe.
    /// </summary>
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
                existing.UpdateLastAccess(); // Reset inactivity timeout
                return existing;
            }

            // Create new handle
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
    public ExcelWorkbookHandle GetHandle(string filePath)
    {
        ObjectDisposedException.ThrowIf(_disposed, nameof(FileHandleManager));

        string absolutePath = Path.GetFullPath(filePath);

        if (_handles.TryGetValue(absolutePath, out var handle))
            return handle;

        throw new InvalidOperationException($"File not open: {filePath}. Call OpenOrGetAsync first.");
    }

    /// <summary>
    /// Checks if a handle exists for the given file path
    /// </summary>
    public bool HasHandle(string filePath)
    {
        string absolutePath = Path.GetFullPath(filePath);
        return _handles.ContainsKey(absolutePath);
    }

    /// <summary>
    /// Explicitly close handle for file path. Removes from cache and disposes.
    /// </summary>
    public async Task CloseAsync(string filePath)
    {
        ObjectDisposedException.ThrowIf(_disposed, nameof(FileHandleManager));

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
    /// Save specific file (must already be open).
    /// </summary>
    public async Task SaveAsync(string filePath)
    {
        ObjectDisposedException.ThrowIf(_disposed, nameof(FileHandleManager));

        var handle = GetHandle(filePath); // Throws if not open
        await handle.SaveAsync();
    }

    /// <summary>
    /// Background cleanup: Close handles inactive for > timeout.
    /// Runs periodically (every minute).
    /// </summary>
    private async Task CleanupInactiveHandlesAsync()
    {
        if (_disposed)
            return;

        var now = DateTime.UtcNow;
        var toRemove = _handles
            .Where(kvp => (now - kvp.Value.LastAccess) > InactivityTimeout)
            .Select(kvp => kvp.Key)
            .ToList();

        foreach (var path in toRemove)
        {
            try
            {
                await CloseAsync(path);
            }
            catch
            {
                // Ignore cleanup errors
            }
        }
    }

    /// <summary>
    /// Gets list of currently open file paths
    /// </summary>
    public IReadOnlyList<string> GetOpenFiles()
    {
        return _handles.Keys.ToList();
    }

    /// <summary>
    /// Disposes the FileHandleManager and all open handles
    /// </summary>
    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;

        // Stop cleanup timer
        _cleanupTimer?.Dispose();

        // Close all handles
        foreach (var handle in _handles.Values)
        {
            try
            {
                handle.Dispose();
            }
            catch
            {
                // Ignore disposal errors
            }
        }

        _handles.Clear();
        _lock.Dispose();
    }
}
