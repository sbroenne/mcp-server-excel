using Sbroenne.ExcelMcp.Core;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests;

/// <summary>
/// Collection fixture for Excel integration tests that enables instance pooling.
/// This dramatically improves test execution speed by reusing Excel instances
/// across tests on the same workbook (~95% faster for cached instances).
/// </summary>
[CollectionDefinition(nameof(ExcelPooledTestCollection))]
public class ExcelPooledTestCollection : ICollectionFixture<ExcelPooledTestFixture>
{
    // This class has no code, and is never created. Its purpose is simply
    // to be the place to apply [CollectionDefinition] and all the
    // ICollectionFixture<> interfaces.
}

/// <summary>
/// Fixture that provides Excel instance pooling for integration tests.
/// Sets up pooling on construction and cleans up on disposal.
/// </summary>
public class ExcelPooledTestFixture : IDisposable
{
    private readonly ExcelInstancePool _pool;

    /// <summary>
    /// Gets the shared Excel instance pool for tests.
    /// Tests should evict instances from the pool before deleting test files.
    /// </summary>
    public ExcelInstancePool Pool => _pool;

    /// <summary>
    /// Safely cleans up a test file by closing its workbook in the pool before deletion.
    /// Call this in test Dispose() methods before deleting files.
    /// </summary>
    /// <param name="filePath">Path to the Excel file to clean up</param>
    public void SafeCleanupFile(string filePath)
    {
        if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            return;

        try
        {
            // Close workbook in pool (if cached)
            _pool.CloseWorkbook(filePath);

            // Small delay to ensure Excel releases file handle
            Thread.Sleep(100);

            // Now safe to delete
            File.Delete(filePath);
        }
        catch
        {
            // Ignore cleanup errors
        }
    }

    public ExcelPooledTestFixture()
    {
        // Create pool with 30-second timeout for tests (shorter than production)
        _pool = new ExcelInstancePool(idleTimeout: TimeSpan.FromSeconds(30));

        // Enable pooling for all Core commands
        ExcelHelper.InstancePool = _pool;
    }

    /// <summary>
    /// Safely clean up a test file by evicting it from the pool first, then deleting it.
    /// This prevents pool corruption when tests delete their Excel files.
    /// </summary>
    /// <param name="filePath">Path to the Excel file to clean up</param>
    public void CleanupTestFile(string filePath)
    {
        try
        {
            // Evict from pool first
            _pool.EvictInstance(filePath);

            // Then delete the file
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }
        catch
        {
            // Ignore cleanup errors
        }
    }

    /// <summary>
    /// Safely clean up a test directory by evicting all Excel files from the pool first, then deleting the directory.
    /// </summary>
    /// <param name="directoryPath">Path to the directory to clean up</param>
    public void CleanupTestDirectory(string directoryPath)
    {
        try
        {
            if (Directory.Exists(directoryPath))
            {
                // Evict all Excel files in the directory from the pool first
                foreach (var file in Directory.GetFiles(directoryPath, "*.xls*", SearchOption.AllDirectories))
                {
                    _pool.EvictInstance(file);
                }

                // Then delete the directory
                Directory.Delete(directoryPath, recursive: true);
            }
        }
        catch
        {
            // Ignore cleanup errors
        }
    }

    public void Dispose()
    {
        // Disable pooling first to prevent new operations from starting
        ExcelHelper.InstancePool = null;

        // Wait for any in-flight operations to complete
        // Excel COM operations can take time, especially during cleanup
        Thread.Sleep(500);

        // Now safe to dispose pool and clean up all Excel instances
        _pool?.Dispose();

        GC.SuppressFinalize(this);
    }
}
