using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.Core;

/// <summary>
/// Manages a pool of Excel COM instances for reuse across multiple operations.
/// Optimized for MCP Server conversational workflows where multiple commands
/// operate on the same workbook in quick succession.
/// </summary>
[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
public sealed class ExcelInstancePool : IDisposable
{
    private readonly ConcurrentDictionary<string, PooledExcelInstance> _instances = new();
    private readonly TimeSpan _idleTimeout;
    private readonly Timer _cleanupTimer;
    private readonly SemaphoreSlim _instanceSemaphore;
    private readonly int _maxInstances;
    private bool _disposed;

    // Metrics
    private long _totalGets;
    private long _totalHits;

    /// <summary>
    /// Gets the current number of active pooled Excel instances.
    /// </summary>
    public int ActiveInstances => _instances.Count;

    /// <summary>
    /// Gets the total number of pool access requests.
    /// </summary>
    public long TotalGets => Interlocked.Read(ref _totalGets);

    /// <summary>
    /// Gets the total number of cache hits (reused instances).
    /// </summary>
    public long TotalHits => Interlocked.Read(ref _totalHits);

    /// <summary>
    /// Gets the cache hit rate (0.0 to 1.0).
    /// </summary>
    public double HitRate => TotalGets > 0 ? (double)TotalHits / TotalGets : 0;

    /// <summary>
    /// Creates a new Excel instance pool with the specified idle timeout and maximum instances.
    /// </summary>
    /// <param name="idleTimeout">Time before idle instances are disposed. Default: 60 seconds.</param>
    /// <param name="maxInstances">Maximum number of concurrent Excel instances. Default: 10.</param>
    public ExcelInstancePool(TimeSpan? idleTimeout = null, int maxInstances = 10)
    {
        _idleTimeout = idleTimeout ?? TimeSpan.FromSeconds(60);
        _maxInstances = maxInstances;
        _instanceSemaphore = new SemaphoreSlim(maxInstances, maxInstances);

        // Cleanup timer runs every 30 seconds to dispose idle instances
        _cleanupTimer = new Timer(CleanupIdleInstances, null,
            TimeSpan.FromSeconds(30), TimeSpan.FromSeconds(30));
    }

    /// <summary>
    /// Executes an action with a pooled Excel instance, reusing existing instance if available.
    /// </summary>
    /// <typeparam name="T">Return type of the action</typeparam>
    /// <param name="filePath">Path to the Excel file</param>
    /// <param name="save">Whether to save changes to the file</param>
    /// <param name="action">Action to execute with Excel application and workbook</param>
    /// <returns>Result of the action</returns>
    /// <exception cref="ExcelPoolCapacityException">Thrown when pool is at maximum capacity and no slot becomes available within timeout period</exception>
    public T WithPooledExcel<T>(string filePath, bool save, Func<dynamic, dynamic, T> action)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        // Track pool access
        Interlocked.Increment(ref _totalGets);

        // Normalize path for pooling (case-insensitive on Windows)
        string normalizedPath = Path.GetFullPath(filePath).ToLowerInvariant();

        // Wait for available slot with timeout (prevents indefinite blocking)
        // Timeout of 5 seconds is reasonable - if pool is full, LLM should know quickly
        if (!_instanceSemaphore.Wait(TimeSpan.FromSeconds(5)))
        {
            // Pool is at capacity - provide actionable guidance to LLM
            throw new ExcelPoolCapacityException(ActiveInstances, _maxInstances, _idleTimeout);
        }

        try
        {
            // Get or create pooled instance
            bool isExistingInstance = _instances.ContainsKey(normalizedPath);
            var pooledInstance = _instances.GetOrAdd(normalizedPath, _ => CreatePooledInstance(filePath));

            // Track cache hit
            if (isExistingInstance)
            {
                Interlocked.Increment(ref _totalHits);
            }

            lock (pooledInstance.Lock)
            {
                try
                {
                    // Update last used timestamp
                    pooledInstance.LastUsed = DateTime.UtcNow;

                    // If workbook is not currently open, open it
                    if (pooledInstance.Workbook == null)
                    {
                        pooledInstance.Workbook = OpenWorkbook(pooledInstance.Excel, filePath);
                    }

                    // Execute the user action
                    T result = action(pooledInstance.Excel, pooledInstance.Workbook);

                    // Save if requested
                    if (save && pooledInstance.Workbook != null)
                    {
                        pooledInstance.Workbook.Save();
                    }

                    return result;
                }
                catch (COMException comEx) when (comEx.ErrorCode == unchecked((int)0x800A03EC))
                {
                    // Excel object is no longer valid - recreate instance
                    DisposePooledInstance(pooledInstance, normalizedPath);

                    // Retry with fresh instance
                    var newInstance = CreatePooledInstance(filePath);
                    _instances[normalizedPath] = newInstance;

                    lock (newInstance.Lock)
                    {
                        newInstance.Workbook = OpenWorkbook(newInstance.Excel, filePath);
                        T result = action(newInstance.Excel, newInstance.Workbook);

                        if (save && newInstance.Workbook != null)
                        {
                            newInstance.Workbook.Save();
                        }

                        return result;
                    }
                }
                catch
                {
                    // On error, close workbook but keep Excel instance alive for retry
                    if (pooledInstance.Workbook != null)
                    {
                        try
                        {
                            pooledInstance.Workbook.Close(false);
                            Marshal.ReleaseComObject(pooledInstance.Workbook);
                            pooledInstance.Workbook = null;
                        }
                        catch (Exception)
                        {
                            // Ignore cleanup errors
                        }
                    }
                    throw;
                }
            }
        }
        finally
        {
            // Always release semaphore slot
            _instanceSemaphore.Release();
        }
    }

    /// <summary>
    /// Closes the workbook for the specified file path, keeping Excel instance alive.
    /// Useful after save operations to release file lock while keeping instance pooled.
    /// </summary>
    /// <param name="filePath">Path to the Excel file</param>
    public void CloseWorkbook(string filePath)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        string normalizedPath = Path.GetFullPath(filePath).ToLowerInvariant();

        if (_instances.TryGetValue(normalizedPath, out var pooledInstance))
        {
            lock (pooledInstance.Lock)
            {
                if (pooledInstance.Workbook != null)
                {
                    try
                    {
                        pooledInstance.Workbook.Close(false);
                        Marshal.ReleaseComObject(pooledInstance.Workbook);
                    }
                    catch
                    {
                        // Ignore errors during close
                    }
                    finally
                    {
                        pooledInstance.Workbook = null;
                    }
                }
            }
        }
    }

    /// <summary>
    /// Removes and disposes the pooled instance for the specified file path.
    /// </summary>
    /// <param name="filePath">Path to the Excel file</param>
    public void EvictInstance(string filePath)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        string normalizedPath = Path.GetFullPath(filePath).ToLowerInvariant();

        if (_instances.TryRemove(normalizedPath, out var pooledInstance))
        {
            DisposePooledInstance(pooledInstance, normalizedPath);

            // Release semaphore slot for evicted instance
            _instanceSemaphore.Release();
        }
    }

    private PooledExcelInstance CreatePooledInstance(string filePath)
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application");
        if (excelType == null)
        {
            throw new InvalidOperationException("Excel is not installed or not properly registered.");
        }

#pragma warning disable IL2072 // COM interop is not AOT compatible but is required for Excel automation
#pragma warning disable CS8600 // Converting null literal to non-nullable type - validated by subsequent null check
        dynamic excel = Activator.CreateInstance(excelType);
#pragma warning restore CS8600
#pragma warning restore IL2072

        if (excel == null)
        {
            throw new InvalidOperationException("Failed to create Excel COM instance.");
        }

        // Configure Excel for automation
        excel.Visible = false;
        excel.DisplayAlerts = false;
        excel.ScreenUpdating = false;
        excel.Interactive = false;

        return new PooledExcelInstance
        {
            Excel = excel,
            Workbook = null,
            LastUsed = DateTime.UtcNow,
            Lock = new object()
        };
    }

    private static dynamic OpenWorkbook(dynamic excel, string filePath)
    {
        string fullPath = Path.GetFullPath(filePath);

        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException($"Excel file not found: {fullPath}", fullPath);
        }

        try
        {
            return excel.Workbooks.Open(fullPath);
        }
        catch (COMException comEx) when (comEx.ErrorCode == unchecked((int)0x8001010A))
        {
            throw new InvalidOperationException(
                "Excel is busy (likely has a dialog open). Close any Excel dialogs and retry.", comEx);
        }
        catch (COMException comEx) when (comEx.ErrorCode == unchecked((int)0x80070020))
        {
            throw new InvalidOperationException(
                $"File '{Path.GetFileName(fullPath)}' is locked by another process.", comEx);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"Failed to open workbook '{Path.GetFileName(fullPath)}'.", ex);
        }
    }

    private void CleanupIdleInstances(object? state)
    {
        if (_disposed) return;

        var now = DateTime.UtcNow;

        // Snapshot keys to avoid modification during enumeration (prevents race condition)
        var keysSnapshot = _instances.Keys.ToList();

        foreach (var key in keysSnapshot)
        {
            if (_instances.TryGetValue(key, out var instance))
            {
                // Check if instance has been idle too long
                if (now - instance.LastUsed > _idleTimeout)
                {
                    // Try to remove and dispose
                    if (_instances.TryRemove(key, out var removedInstance))
                    {
                        DisposePooledInstance(removedInstance, key);

                        // Release semaphore slot for removed instance
                        _instanceSemaphore.Release();
                    }
                }
            }
        }
    }

    private static void DisposePooledInstance(PooledExcelInstance instance, string path)
    {
        lock (instance.Lock)
        {
            // Close workbook
            if (instance.Workbook != null)
            {
                try
                {
                    instance.Workbook.Close(false);
                    Marshal.ReleaseComObject(instance.Workbook);
                }
                catch
                {
                    // Ignore errors during cleanup
                }
                finally
                {
                    instance.Workbook = null;
                }
            }

            // Quit Excel
            if (instance.Excel != null)
            {
                try
                {
                    instance.Excel.Quit();
                    Marshal.ReleaseComObject(instance.Excel);
                }
                catch
                {
                    // Ignore errors during cleanup
                }
                finally
                {
                    instance.Excel = null;
                }
            }
        }

        // Force GC to clean up COM objects
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }

    /// <summary>
    /// Disposes all pooled Excel instances and releases resources.
    /// </summary>
    public void Dispose()
    {
        if (_disposed) return;

        _disposed = true;
        _cleanupTimer?.Dispose();

        // Dispose all pooled instances
        foreach (var kvp in _instances)
        {
            DisposePooledInstance(kvp.Value, kvp.Key);
        }

        _instances.Clear();

        // Dispose semaphore
        _instanceSemaphore?.Dispose();
    }

    private class PooledExcelInstance
    {
        public dynamic? Excel { get; set; }
        public dynamic? Workbook { get; set; }
        public DateTime LastUsed { get; set; }
        public required object Lock { get; set; }
    }
}
