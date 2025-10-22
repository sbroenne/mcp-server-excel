using Sbroenne.ExcelMcp.Core;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Static manager for Excel instance pool used by MCP Server tools.
/// Provides thread-safe singleton access to the pool for all static tool methods.
/// </summary>
public static class ExcelToolsPoolManager
{
    private static ExcelInstancePool? _pool;
    private static readonly object _lock = new();

    /// <summary>
    /// Gets the current Excel instance pool. Returns null if not initialized.
    /// </summary>
    public static ExcelInstancePool? Pool
    {
        get
        {
            lock (_lock)
            {
                return _pool;
            }
        }
    }

    /// <summary>
    /// Gets whether pooling is enabled (pool is initialized and available).
    /// </summary>
    public static bool IsPoolingEnabled => Pool != null;

    /// <summary>
    /// Initializes the Excel instance pool for use by all MCP tools.
    /// Should be called once during application startup.
    /// </summary>
    /// <param name="pool">The Excel instance pool to use</param>
    public static void Initialize(ExcelInstancePool pool)
    {
        lock (_lock)
        {
            if (_pool != null)
            {
                throw new InvalidOperationException("Excel instance pool is already initialized");
            }
            _pool = pool ?? throw new ArgumentNullException(nameof(pool));
        }
    }

    /// <summary>
    /// Shuts down the Excel instance pool and disposes all pooled instances.
    /// Should be called during application shutdown.
    /// </summary>
    public static void Shutdown()
    {
        lock (_lock)
        {
            if (_pool != null)
            {
                _pool.Dispose();
                _pool = null;
            }
        }
    }

    /// <summary>
    /// Executes an action with pooled Excel if available, falls back to WithExcel if not.
    /// This provides backward compatibility and graceful degradation.
    /// </summary>
    /// <typeparam name="T">Return type of the action</typeparam>
    /// <param name="filePath">Path to the Excel file</param>
    /// <param name="save">Whether to save changes to the file</param>
    /// <param name="action">Action to execute with Excel application and workbook</param>
    /// <returns>Result of the action</returns>
    public static T WithExcel<T>(string filePath, bool save, Func<dynamic, dynamic, T> action)
    {
        var pool = Pool;

        // Use pooled instance for better performance if available; otherwise, fall back to single-instance pattern.
        return pool != null
            ? pool.WithPooledExcel(filePath, save, action)
            : ExcelHelper.WithExcel(filePath, save, action);
    }
}
