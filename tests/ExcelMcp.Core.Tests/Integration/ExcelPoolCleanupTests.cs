using System.Diagnostics;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration;

/// <summary>
/// Focused tests for Excel instance pool cleanup verification.
/// Tests verify that Excel processes are properly cleaned up after pool disposal.
/// These tests are marked as "OnDemand" and don't run by default - use explicit filter to run them.
/// Run with: dotnet test --filter "RunType=OnDemand"
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Core")]
[Trait("Feature", "PoolCleanup")]
[Trait("RequiresExcel", "true")]
[Trait("RunType", "OnDemand")]  // Don't run by default - only when explicitly filtered
public class ExcelPoolCleanupTests
{
    private readonly ITestOutputHelper _output;

    public ExcelPoolCleanupTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public void PoolDisposal_WithMultipleInstances_ShouldCleanupAllExcelProcesses()
    {
        // Get baseline Excel process count
        int initialExcelCount = GetExcelProcessCount();
        _output.WriteLine($"Initial Excel processes: {initialExcelCount}");

        // Create test files
        var testFiles = new List<string>();
        var fileCommands = new FileCommands();

        for (int i = 0; i < 3; i++)
        {
            var testFile = Path.Combine(Path.GetTempPath(), $"pool_cleanup_test_{i}_{Guid.NewGuid()}.xlsx");
            fileCommands.CreateEmpty(testFile, overwriteIfExists: true);
            testFiles.Add(testFile);
        }

        try
        {
            // Create pool and perform operations
            var pool = new ExcelInstancePool(
                idleTimeout: TimeSpan.FromSeconds(30),
                maxInstances: 10
            );

            // Perform operations on each file to create pooled instances
            foreach (var testFile in testFiles)
            {
                pool.WithPooledExcel(testFile, save: false, (excel, workbook) =>
                {
                    // Simple operation to ensure Excel instance is created
                    dynamic sheets = workbook.Worksheets;
                    int count = sheets.Count;
                    return count;
                });
            }

            // Check that Excel processes increased
            int duringPoolCount = GetExcelProcessCount();
            _output.WriteLine($"Excel processes during pool: {duringPoolCount}");
            _output.WriteLine($"Active instances in pool: {pool.ActiveInstances}");

            Assert.True(duringPoolCount > initialExcelCount,
                $"Expected Excel processes to increase. Initial: {initialExcelCount}, During: {duringPoolCount}");

            // Dispose pool and verify cleanup
            _output.WriteLine("Disposing pool...");
            pool.Dispose();

            // Wait for Excel processes to terminate
            _output.WriteLine("Waiting for Excel cleanup (3 seconds)...");
            Thread.Sleep(3000);

            // Force GC to clean up any remaining COM references
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            // Check final Excel process count
            int finalExcelCount = GetExcelProcessCount();
            _output.WriteLine($"Final Excel processes: {finalExcelCount}");

            // Assert that we're back to baseline (or very close)
            int excelDifference = finalExcelCount - initialExcelCount;
            Assert.True(excelDifference <= 1,
                $"Expected Excel process count to return to baseline. Initial: {initialExcelCount}, Final: {finalExcelCount}, Difference: {excelDifference}");
        }
        finally
        {
            // Cleanup test files
            foreach (var testFile in testFiles)
            {
                try
                {
                    if (File.Exists(testFile))
                    {
                        File.Delete(testFile);
                    }
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }
    }

    [Fact]
    public void PoolDisposal_WithEviction_ShouldCleanupImmediately()
    {
        // Get baseline Excel process count
        int initialExcelCount = GetExcelProcessCount();
        _output.WriteLine($"Initial Excel processes: {initialExcelCount}");

        // Create test file
        var fileCommands = new FileCommands();
        var testFile = Path.Combine(Path.GetTempPath(), $"pool_eviction_test_{Guid.NewGuid()}.xlsx");
        fileCommands.CreateEmpty(testFile, overwriteIfExists: true);

        try
        {
            // Create pool
            var pool = new ExcelInstancePool(
                idleTimeout: TimeSpan.FromSeconds(30),
                maxInstances: 10
            );

            // Perform operation to create pooled instance
            pool.WithPooledExcel(testFile, save: false, (excel, workbook) =>
            {
                dynamic sheets = workbook.Worksheets;
                return sheets.Count;
            });

            int afterOperationCount = GetExcelProcessCount();
            _output.WriteLine($"Excel processes after operation: {afterOperationCount}");

            // Evict the instance
            _output.WriteLine("Evicting instance...");
            pool.EvictInstance(testFile);

            // Wait briefly for eviction
            Thread.Sleep(1000);

            int afterEvictionCount = GetExcelProcessCount();
            _output.WriteLine($"Excel processes after eviction: {afterEvictionCount}");

            // Dispose pool
            _output.WriteLine("Disposing pool...");
            pool.Dispose();

            // Wait for cleanup
            Thread.Sleep(2000);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            int finalExcelCount = GetExcelProcessCount();
            _output.WriteLine($"Final Excel processes: {finalExcelCount}");

            // Assert cleanup
            int excelDifference = finalExcelCount - initialExcelCount;
            Assert.True(excelDifference <= 1,
                $"Expected Excel process count to return to baseline. Initial: {initialExcelCount}, Final: {finalExcelCount}, Difference: {excelDifference}");
        }
        finally
        {
            // Cleanup test file
            try
            {
                if (File.Exists(testFile))
                {
                    File.Delete(testFile);
                }
            }
            catch
            {
                // Ignore cleanup errors
            }
        }
    }

    [Fact]
    public void PoolDisposal_WithMultipleFilesUsed_ShouldCleanupAllProcesses()
    {
        // Get baseline Excel process count
        int initialExcelCount = GetExcelProcessCount();
        _output.WriteLine($"Initial Excel processes: {initialExcelCount}");

        var fileCommands = new FileCommands();
        var testFiles = new List<string>();

        try
        {
            // Create pool
            var pool = new ExcelInstancePool(
                idleTimeout: TimeSpan.FromSeconds(30),
                maxInstances: 10
            );

            // Create and use multiple test files
            for (int i = 0; i < 5; i++)
            {
                var testFile = Path.Combine(Path.GetTempPath(), $"multi_file_test_{i}_{Guid.NewGuid()}.xlsx");
                fileCommands.CreateEmpty(testFile, overwriteIfExists: true);
                testFiles.Add(testFile);

                // Use pool for each file
                pool.WithPooledExcel(testFile, save: false, (excel, workbook) =>
                {
                    dynamic sheets = workbook.Worksheets;
                    return sheets.Count;
                });
            }

            int duringTestsCount = GetExcelProcessCount();
            _output.WriteLine($"Excel processes during tests: {duringTestsCount}");
            _output.WriteLine($"Active instances in pool: {pool.ActiveInstances}");

            // Dispose pool
            _output.WriteLine("Disposing pool...");
            pool.Dispose();

            // Wait for cleanup
            _output.WriteLine("Waiting for Excel cleanup (3 seconds)...");
            Thread.Sleep(3000);

            // Force GC
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            int finalExcelCount = GetExcelProcessCount();
            _output.WriteLine($"Final Excel processes: {finalExcelCount}");

            // Assert cleanup
            int excelDifference = finalExcelCount - initialExcelCount;
            _output.WriteLine($"Excel process difference from baseline: {excelDifference}");

            Assert.True(excelDifference <= 1,
                $"Expected Excel process count to return close to baseline. Initial: {initialExcelCount}, Final: {finalExcelCount}, Difference: {excelDifference}");
        }
        finally
        {
            // Cleanup test files
            foreach (var testFile in testFiles)
            {
                try
                {
                    if (File.Exists(testFile))
                    {
                        File.Delete(testFile);
                    }
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }
    }

    [Fact]
    public void StressTest_ParallelOperationsOnManyFiles_ShouldCleanupAllProcesses()
    {
        // Get baseline Excel process count
        int initialExcelCount = GetExcelProcessCount();
        _output.WriteLine($"=== STRESS TEST: Parallel Operations ===");
        _output.WriteLine($"Initial Excel processes: {initialExcelCount}");

        var pool = new ExcelInstancePool(
            idleTimeout: TimeSpan.FromSeconds(30),
            maxInstances: 10
        );

        var fileCommands = new FileCommands();
        var testFiles = new List<string>();

        try
        {
            // Create 20 test files (simulating many integration tests)
            _output.WriteLine("Creating 20 test files...");
            for (int i = 0; i < 20; i++)
            {
                var testFile = Path.Combine(Path.GetTempPath(), $"stress_test_{i}_{Guid.NewGuid()}.xlsx");
                fileCommands.CreateEmpty(testFile, overwriteIfExists: true);
                testFiles.Add(testFile);
            }

            // Perform multiple operations on random files in parallel (simulating concurrent tests)
            _output.WriteLine("Performing 50 parallel operations across files...");
            var tasks = new List<Task>();

            for (int i = 0; i < 50; i++)
            {
                var iteration = i;
                var fileIndex = i % testFiles.Count; // Distribute operations across files
                var task = Task.Run(() =>
                {
                    try
                    {
                        // Pick a file (cycling through them)
                        var testFile = testFiles[fileIndex];

                        // Perform operation using pool directly
                        pool.WithPooledExcel(testFile, save: false, (excel, workbook) =>
                        {
                            dynamic sheets = workbook.Worksheets;
                            int count = sheets.Count;

                            // Simulate some work
                            Thread.Sleep((iteration % 5) * 10 + 10);

                            return count;
                        });
                    }
                    catch (Exception ex)
                    {
                        _output.WriteLine($"Operation {iteration} failed: {ex.Message}");
                    }
                });

                tasks.Add(task);
            }

            // Wait for all operations to complete
#pragma warning disable xUnit1031 // Test methods should not use blocking task operations - intentional for stress test
            Task.WhenAll(tasks).GetAwaiter().GetResult();
#pragma warning restore xUnit1031
            _output.WriteLine("All operations completed");

            // Check state during operations
            int duringTestsCount = GetExcelProcessCount();
            int activeInstances = pool.ActiveInstances;
            long hitRate = pool.TotalHits;

            _output.WriteLine($"Excel processes during tests: {duringTestsCount}");
            _output.WriteLine($"Active instances in pool: {activeInstances}");
            _output.WriteLine($"Total cache hits: {hitRate}");
            _output.WriteLine($"Hit rate: {pool.HitRate:P}");

            // Access some files again to test reuse
            _output.WriteLine("Accessing files again to test reuse...");
            foreach (var testFile in testFiles.Take(5))
            {
                pool.WithPooledExcel(testFile, save: false, (excel, workbook) =>
                {
                    dynamic sheets = workbook.Worksheets;
                    return sheets.Count;
                });
            }

            int afterReuseCount = GetExcelProcessCount();
            _output.WriteLine($"Excel processes after reuse: {afterReuseCount}");

            // Dispose pool
            _output.WriteLine("Disposing pool...");
            pool.Dispose();

            // Wait for cleanup (longer for stress test)
            _output.WriteLine("Waiting for Excel cleanup (5 seconds)...");
            Thread.Sleep(5000);

            // Force GC multiple times
            for (int i = 0; i < 3; i++)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            GC.Collect();

            int finalExcelCount = GetExcelProcessCount();
            _output.WriteLine($"Final Excel processes: {finalExcelCount}");

            // Assert cleanup
            int excelDifference = finalExcelCount - initialExcelCount;
            _output.WriteLine($"Excel process difference from baseline: {excelDifference}");

            Assert.True(excelDifference <= 2,
                $"Expected Excel process count to return close to baseline after stress test. Initial: {initialExcelCount}, Final: {finalExcelCount}, Difference: {excelDifference}");
        }
        finally
        {
            // Cleanup test files
            _output.WriteLine("Cleaning up test files...");
            foreach (var testFile in testFiles)
            {
                try
                {
                    if (File.Exists(testFile))
                    {
                        File.Delete(testFile);
                    }
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }
    }

    [Fact]
    public void StressTest_RapidCreateAndDispose_ShouldNotLeakProcesses()
    {
        // Get baseline Excel process count
        int initialExcelCount = GetExcelProcessCount();
        _output.WriteLine($"=== STRESS TEST: Rapid Create/Dispose ===");
        _output.WriteLine($"Initial Excel processes: {initialExcelCount}");

        var fileCommands = new FileCommands();
        var allTestFiles = new List<string>();

        try
        {
            // Simulate multiple test runs (like running tests multiple times)
            for (int run = 0; run < 3; run++)
            {
                _output.WriteLine($"\n--- Test Run {run + 1} ---");

                var pool = new ExcelInstancePool(
                    idleTimeout: TimeSpan.FromSeconds(30),
                    maxInstances: 10
                );

                var testFiles = new List<string>();

                // Create files
                for (int i = 0; i < 10; i++)
                {
                    var testFile = Path.Combine(Path.GetTempPath(), $"rapid_test_run{run}_{i}_{Guid.NewGuid()}.xlsx");
                    fileCommands.CreateEmpty(testFile, overwriteIfExists: true);
                    testFiles.Add(testFile);
                    allTestFiles.Add(testFile);
                }

                // Perform operations
                foreach (var testFile in testFiles)
                {
                    pool.WithPooledExcel(testFile, save: false, (excel, workbook) =>
                    {
                        dynamic sheets = workbook.Worksheets;
                        return sheets.Count;
                    });
                }

                int duringCount = GetExcelProcessCount();
                _output.WriteLine($"Excel processes during run: {duringCount}");
                _output.WriteLine($"Active instances: {pool.ActiveInstances}");

                // Dispose pool immediately (simulating test fixture disposal)
                Thread.Sleep(300);
                pool.Dispose();
                Thread.Sleep(1000);

                int afterRunCount = GetExcelProcessCount();
                _output.WriteLine($"Excel processes after disposal: {afterRunCount}");

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }

            // Final cleanup check
            _output.WriteLine("\n--- Final Cleanup Check ---");
            Thread.Sleep(3000);

            for (int i = 0; i < 3; i++)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            GC.Collect();

            int finalExcelCount = GetExcelProcessCount();
            _output.WriteLine($"Final Excel processes: {finalExcelCount}");

            // Assert cleanup
            int excelDifference = finalExcelCount - initialExcelCount;
            _output.WriteLine($"Excel process difference from baseline: {excelDifference}");

            Assert.True(excelDifference <= 2,
                $"Expected Excel process count to return close to baseline after rapid create/dispose. Initial: {initialExcelCount}, Final: {finalExcelCount}, Difference: {excelDifference}");
        }
        finally
        {
            // Cleanup all test files
            _output.WriteLine("Cleaning up all test files...");
            foreach (var testFile in allTestFiles)
            {
                try
                {
                    if (File.Exists(testFile))
                    {
                        File.Delete(testFile);
                    }
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }
    }

    private int GetExcelProcessCount()
    {
        try
        {
            var processes = Process.GetProcessesByName("excel");
            int count = processes.Length;

            // Release process handles
            foreach (var process in processes)
            {
                process.Dispose();
            }

            return count;
        }
        catch
        {
            return 0;
        }
    }
}
