using System.Diagnostics;
using Sbroenne.ExcelMcp.Core;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration;

/// <summary>
/// Integration tests for Excel instance pooling that verify actual reuse behavior,
/// pool metrics, and Excel process management.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "Pooling")]
[Trait("RequiresExcel", "true")]
public class ExcelInstancePoolIntegrationTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _testFile1;
    private readonly string _testFile2;
    private readonly string _testFile3;
    private readonly ExcelInstancePool _pool;
    private readonly FileCommands _fileCommands;

    public ExcelInstancePoolIntegrationTests(ITestOutputHelper output)
    {
        _output = output;
        _fileCommands = new FileCommands();

        // Create test files
        _testFile1 = Path.Combine(Path.GetTempPath(), $"pool_integration_1_{Guid.NewGuid()}.xlsx");
        _testFile2 = Path.Combine(Path.GetTempPath(), $"pool_integration_2_{Guid.NewGuid()}.xlsx");
        _testFile3 = Path.Combine(Path.GetTempPath(), $"pool_integration_3_{Guid.NewGuid()}.xlsx");

        _fileCommands.CreateEmpty(_testFile1, overwriteIfExists: true);
        _fileCommands.CreateEmpty(_testFile2, overwriteIfExists: true);
        _fileCommands.CreateEmpty(_testFile3, overwriteIfExists: true);

        // Create pool with reasonable settings for testing
        _pool = new ExcelInstancePool(
            idleTimeout: TimeSpan.FromSeconds(30),
            maxInstances: 10
        );
    }

    [Fact]
    public void PoolMetrics_MultipleOperations_ShouldShowReuse()
    {
        _output.WriteLine("=== Testing Pool Reuse Metrics ===");

        // Initial state
        _output.WriteLine($"Initial - Active: {_pool.ActiveInstances}, Total Hits: {_pool.TotalHits}, Hit Rate: {_pool.HitRate:P2}");
        Assert.Equal(0, _pool.ActiveInstances);
        Assert.Equal(0, _pool.TotalHits);

        // First operation - creates new instance (MISS, not hit)
        _pool.WithPooledExcel(_testFile1, false, (excel, workbook) => "Op1");
        _output.WriteLine($"After Op1 - Active: {_pool.ActiveInstances}, Total Hits: {_pool.TotalHits}, Hit Rate: {_pool.HitRate:P2}");
        Assert.Equal(1, _pool.ActiveInstances);
        Assert.Equal(0, _pool.TotalHits); // First operation is MISS, not hit

        // Second operation on same file - should reuse (HIT)
        _pool.WithPooledExcel(_testFile1, false, (excel, workbook) => "Op2");
        _output.WriteLine($"After Op2 - Active: {_pool.ActiveInstances}, Total Hits: {_pool.TotalHits}, Hit Rate: {_pool.HitRate:P2}");
        Assert.Equal(1, _pool.ActiveInstances); // Still 1 instance
        Assert.Equal(1, _pool.TotalHits); // First HIT (reused existing)

        // Third operation on same file - should reuse again (HIT #2)
        _pool.WithPooledExcel(_testFile1, false, (excel, workbook) => "Op3");
        _output.WriteLine($"After Op3 - Active: {_pool.ActiveInstances}, Total Hits: {_pool.TotalHits}, Hit Rate: {_pool.HitRate:P2}");
        Assert.Equal(1, _pool.ActiveInstances); // Still 1 instance
        Assert.Equal(2, _pool.TotalHits); // Second HIT

        // Verify hit rate (2 hits out of 3 total operations = 66.67%)
        Assert.Equal(2.0 / 3.0, _pool.HitRate, precision: 2);
    }

    [Fact]
    public void PoolMetrics_MultipleFiles_ShouldCreateMultipleInstances()
    {
        _output.WriteLine("=== Testing Multiple File Pooling ===");

        // Open file 1
        _pool.WithPooledExcel(_testFile1, false, (excel, workbook) => "File1-Op1");
        _output.WriteLine($"After File1 - Active: {_pool.ActiveInstances}");
        Assert.Equal(1, _pool.ActiveInstances);

        // Open file 2 - creates second instance
        _pool.WithPooledExcel(_testFile2, false, (excel, workbook) => "File2-Op1");
        _output.WriteLine($"After File2 - Active: {_pool.ActiveInstances}");
        Assert.Equal(2, _pool.ActiveInstances);

        // Open file 3 - creates third instance
        _pool.WithPooledExcel(_testFile3, false, (excel, workbook) => "File3-Op1");
        _output.WriteLine($"After File3 - Active: {_pool.ActiveInstances}");
        Assert.Equal(3, _pool.ActiveInstances);

        // Reuse file 1 - should not create new instance
        _pool.WithPooledExcel(_testFile1, false, (excel, workbook) => "File1-Op2");
        _output.WriteLine($"After File1 reuse - Active: {_pool.ActiveInstances}");
        Assert.Equal(3, _pool.ActiveInstances); // Still 3

        // Reuse file 2 - should not create new instance
        _pool.WithPooledExcel(_testFile2, false, (excel, workbook) => "File2-Op2");
        _output.WriteLine($"After File2 reuse - Active: {_pool.ActiveInstances}");
        Assert.Equal(3, _pool.ActiveInstances); // Still 3

        // Total operations: 5 (3 creates + 2 reuses)
        // Expected hits: 2 (File1 reuse + File2 reuse)
        _output.WriteLine($"Final - Total Hits: {_pool.TotalHits}, Active: {_pool.ActiveInstances}");
        Assert.Equal(2, _pool.TotalHits); // 2 cache hits (reuses)
        Assert.Equal(3, _pool.ActiveInstances); // 3 unique files

        // Hit rate should be 2/5 = 40%
        Assert.Equal(2.0 / 5.0, _pool.HitRate, precision: 2);
    }

    [Fact(Skip = "Excel process counts don't reliably map to pool instances - Windows/Excel COM manages process lifecycle")]
    public void ExcelProcessCount_WithPooling_ShouldMatchActiveInstances()
    {
        _output.WriteLine("=== Testing Excel Process Count Matches Pool Metrics ===");

        // Note: We don't kill all Excel processes - that would close user's personal Excel files!
        // Tests rely on proper pool cleanup via Dispose()

        int initialExcelCount = GetExcelProcessCount();
        _output.WriteLine($"Initial Excel processes: {initialExcelCount}");

        // Create 3 pooled instances
        _pool.WithPooledExcel(_testFile1, false, (excel, workbook) => "File1");
        Thread.Sleep(500); // Allow process to stabilize
        int count1 = GetExcelProcessCount();
        _output.WriteLine($"After File1 - Excel processes: {count1}, Pool active: {_pool.ActiveInstances}");

        _pool.WithPooledExcel(_testFile2, false, (excel, workbook) => "File2");
        Thread.Sleep(500);
        int count2 = GetExcelProcessCount();
        _output.WriteLine($"After File2 - Excel processes: {count2}, Pool active: {_pool.ActiveInstances}");

        _pool.WithPooledExcel(_testFile3, false, (excel, workbook) => "File3");
        Thread.Sleep(500);
        int count3 = GetExcelProcessCount();
        _output.WriteLine($"After File3 - Excel processes: {count3}, Pool active: {_pool.ActiveInstances}");

        // Verify pool metrics match process count
        Assert.Equal(3, _pool.ActiveInstances);

        // Excel process count should match active instances (+/- initial count)
        int expectedProcesses = initialExcelCount + 3;
        _output.WriteLine($"Expected Excel processes: {expectedProcesses}, Actual: {count3}");
        Assert.InRange(count3, expectedProcesses - 1, expectedProcesses + 1);

        // Reuse file1 - should not create new Excel process
        _pool.WithPooledExcel(_testFile1, false, (excel, workbook) => "File1-Reuse");
        Thread.Sleep(500);
        int count4 = GetExcelProcessCount();
        _output.WriteLine($"After File1 reuse - Excel processes: {count4}, Pool active: {_pool.ActiveInstances}");

        Assert.Equal(3, _pool.ActiveInstances); // Still 3 pooled instances
        Assert.InRange(count4, expectedProcesses - 1, expectedProcesses + 1); // Same process count
    }

    [Fact(Skip = "Excel process counts don't reliably map to pool instances - Windows/Excel COM manages process lifecycle")]
    public void ExcelProcessCount_WithEviction_ShouldDecrease()
    {
        _output.WriteLine("=== Testing Excel Process Cleanup on Eviction ===");

        // Note: We don't kill all Excel processes - that would close user's personal Excel files!
        int initialCount = GetExcelProcessCount();
        _output.WriteLine($"Initial Excel processes: {initialCount}");

        // Create 2 pooled instances
        _pool.WithPooledExcel(_testFile1, false, (excel, workbook) => "File1");
        _pool.WithPooledExcel(_testFile2, false, (excel, workbook) => "File2");
        Thread.Sleep(1000); // Wait for processes to stabilize

        int afterPooling = GetExcelProcessCount();
        _output.WriteLine($"After pooling 2 files - Excel processes: {afterPooling}, Pool active: {_pool.ActiveInstances}");
        Assert.Equal(2, _pool.ActiveInstances);

        // Evict file1 - should dispose Excel instance
        _pool.EvictInstance(_testFile1);
        Thread.Sleep(2000); // Wait for Excel COM cleanup (can take 2-5 seconds)

        int afterEviction = GetExcelProcessCount();
        _output.WriteLine($"After evicting File1 - Excel processes: {afterEviction}, Pool active: {_pool.ActiveInstances}");
        Assert.Equal(1, _pool.ActiveInstances);

        // Process count should have decreased
        Assert.True(afterEviction < afterPooling,
            $"Expected fewer Excel processes after eviction. Before: {afterPooling}, After: {afterEviction}");
    }

    [Fact(Skip = "Excel process counts don't reliably map to pool instances - Windows/Excel COM manages process lifecycle")]
    public void ExcelProcessCount_WithDispose_ShouldCleanupAll()
    {
        _output.WriteLine("=== Testing Complete Cleanup on Pool Disposal ===");

        // Note: We don't kill all Excel processes - that would close user's personal Excel files!
        int initialCount = GetExcelProcessCount();
        _output.WriteLine($"Initial Excel processes: {initialCount}");

        // Create a new pool for this test
        using (var testPool = new ExcelInstancePool(TimeSpan.FromSeconds(30), maxInstances: 5))
        {
            // Create 3 pooled instances
            testPool.WithPooledExcel(_testFile1, false, (excel, workbook) => "File1");
            testPool.WithPooledExcel(_testFile2, false, (excel, workbook) => "File2");
            testPool.WithPooledExcel(_testFile3, false, (excel, workbook) => "File3");
            Thread.Sleep(1000);

            int duringPooling = GetExcelProcessCount();
            _output.WriteLine($"During pooling - Excel processes: {duringPooling}, Pool active: {testPool.ActiveInstances}");
            Assert.Equal(3, testPool.ActiveInstances);
        } // Pool disposed here

        Thread.Sleep(3000); // Wait for Excel COM cleanup (disposal of 3 instances can take time)

        int afterDispose = GetExcelProcessCount();
        _output.WriteLine($"After pool disposal - Excel processes: {afterDispose}");

        // All Excel instances should be cleaned up
        Assert.InRange(afterDispose, initialCount, initialCount + 1);
    }

    [Fact]
    public void PoolCapacity_ExceedingLimit_ShouldWaitOrFail()
    {
        _output.WriteLine("=== Testing Pool Capacity Limits ===");

        // Create pool with low capacity
        using (var limitedPool = new ExcelInstancePool(TimeSpan.FromSeconds(30), maxInstances: 2))
        {
            // Fill pool to capacity
            limitedPool.WithPooledExcel(_testFile1, false, (excel, workbook) => "File1");
            limitedPool.WithPooledExcel(_testFile2, false, (excel, workbook) => "File2");

            _output.WriteLine($"Pool at capacity - Active: {limitedPool.ActiveInstances}");
            Assert.Equal(2, limitedPool.ActiveInstances);

            // Attempt to exceed capacity should timeout
            var stopwatch = Stopwatch.StartNew();

            try
            {
                limitedPool.WithPooledExcel(_testFile3, false, (excel, workbook) => "File3");
                Assert.Fail("Expected ExcelPoolCapacityException but operation succeeded");
            }
            catch (ExcelPoolCapacityException ex)
            {
                stopwatch.Stop();
                _output.WriteLine($"Correctly threw ExcelPoolCapacityException after {stopwatch.ElapsedMilliseconds}ms");
                _output.WriteLine($"Exception message: {ex.Message}");

                // Should throw quickly (immediate check) or timeout around 5 seconds (semaphore wait)
                // Microsoft.Extensions.ObjectPool implementation throws immediately (0-100ms) - which is better UX
                Assert.InRange(stopwatch.ElapsedMilliseconds, 0, 7000);

                // Verify exception has helpful message
                Assert.Contains("maximum capacity", ex.Message);
                Assert.Contains("2/2", ex.Message); // Shows capacity
            }
        }
    }

    private int GetExcelProcessCount()
    {
        try
        {
            var processes = Process.GetProcessesByName("excel");
            int count = processes.Length;

            // Dispose process objects
            foreach (var p in processes)
            {
                p.Dispose();
            }

            return count;
        }
        catch
        {
            return 0;
        }
    }

    // KillExcelProcesses() method removed - it's dangerous to kill ALL Excel processes
    // Tests now use baseline process counting instead of forcing zero processes

    public void Dispose()
    {
        _output.WriteLine("=== Test Cleanup ===");

        // Dispose pool
        _pool?.Dispose();
        Thread.Sleep(1000); // Wait for Excel processes to close

        // Delete test files
        foreach (var file in new[] { _testFile1, _testFile2, _testFile3 })
        {
            if (File.Exists(file))
            {
                try
                {
                    File.Delete(file);
                }
                catch (Exception ex)
                {
                    _output.WriteLine($"Failed to delete {file}: {ex.Message}");
                }
            }
        }

        GC.SuppressFinalize(this);
    }
}
