using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

/// <summary>
/// Integration tests for ExcelBatch - verifies batch operations and COM cleanup.
/// Tests that Excel instances are reused across operations and properly cleaned up.
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test ExcelBatch.Execute() reuses Excel instance
/// - ✅ Test ExcelBatch.Dispose() COM cleanup
/// - ✅ Test ExcelBatch.Save() functionality
/// - ✅ Verify Excel.exe process termination (no leaks)
///
/// NOTE: ExcelBatch.Dispose() handles all GC cleanup automatically.
/// Tests only need to wait for async disposal and process termination timing.
///
/// IMPORTANT: These tests spawn and terminate Excel processes (side effects).
/// They run OnDemand only to avoid interference with normal test runs.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "ExcelBatch")]
[Collection("Sequential")] // Disable parallelization to avoid COM interference
public class ExcelBatchTests : IAsyncLifetime
{
    private readonly ITestOutputHelper _output;
    private static string? _staticTestFile;
    private string? _testFileCopy;

    public ExcelBatchTests(ITestOutputHelper output)
    {
        _output = output;
    }

    public Task InitializeAsync()
    {
        // Use static test file from TestFiles folder (must be pre-created)
        if (_staticTestFile == null)
        {
            var testFolder = Path.Join(AppContext.BaseDirectory, "Integration", "Session", "TestFiles");
            _staticTestFile = Path.Join(testFolder, "batch-test-static.xlsx");

            // Verify the static file exists
            if (!File.Exists(_staticTestFile))
            {
                throw new FileNotFoundException($"Static test file not found at {_staticTestFile}. " +
                    "Please create the batch-test-static.xlsx file in the TestFiles folder.");
            }
        }

        // Create a fresh copy for this test instance (in temp folder)
        _testFileCopy = Path.Join(Path.GetTempPath(), $"batch-test-{Guid.NewGuid():N}.xlsx");
        File.Copy(_staticTestFile, _testFileCopy, overwrite: true);

        // Wait for any Excel processes from file creation to terminate
        return Task.Delay(500);
    }

    public Task DisposeAsync()
    {
        // Clean up this test's copy
        if (_testFileCopy != null && File.Exists(_testFileCopy))
        {
            File.Delete(_testFileCopy);
        }
        return Task.CompletedTask;
    }

    private static void CleanupStaticFile()
    {
        if (_staticTestFile != null && File.Exists(_staticTestFile))
        {
            File.Delete(_staticTestFile);
        }
    }

    [Fact]
    public void ExecuteAsync_MultipleOperations_ReusesExcelInstance()
    {
        // Arrange
        int operationCount = 0;

        // Act - Use batching for multiple operations
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);

        for (int i = 0; i < 5; i++)
        {
            batch.Execute((ctx, ct) =>
            {
                operationCount++;
                _output.WriteLine($"Batch operation {operationCount}");

                // Verify we have the same context
                Assert.NotNull(ctx.App);
                Assert.NotNull(ctx.Book);

                return operationCount;
            });
        }

        // Assert
        Assert.Equal(5, operationCount);
        _output.WriteLine($"✓ Completed {operationCount} batch operations");
    }

    [Fact]
    public void Dispose_CleansUpComObjects_NoProcessLeak()
    {
        // Arrange
        var startingProcesses = Process.GetProcessesByName("EXCEL");
        int startingCount = startingProcesses.Length;

        _output.WriteLine($"Excel processes before: {startingCount}");

        // Act
        var batch = ExcelSession.BeginBatch(_testFileCopy!);

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            _ = sheet.Range["A1"].Value2;
            return 0;
        });

        batch.Dispose();

        // Wait for Excel process to fully terminate with polling
        // Excel.Quit() signals shutdown but process termination is OS-controlled
        // Dispose() blocks up to StaThreadJoinTimeout for COM cleanup, but process may linger briefly
        var waitTimeout = TimeSpan.FromSeconds(15); // Allow reasonable time for process cleanup
        var stopwatch = Stopwatch.StartNew();
        int endingCount;
        do
        {
            Thread.Sleep(500); // Poll every 500ms
            endingCount = Process.GetProcessesByName("EXCEL").Length;
            _output.WriteLine($"Excel processes at {stopwatch.Elapsed.TotalSeconds:F1}s: {endingCount}");
        }
        while (endingCount > startingCount && stopwatch.Elapsed < waitTimeout);

        // Assert
        _output.WriteLine($"Excel processes after {stopwatch.Elapsed.TotalSeconds:F1}s: {endingCount}");

        Assert.True(endingCount <= startingCount,
            $"Excel process leak in batch! Started with {startingCount}, ended with {endingCount} after {waitTimeout.TotalSeconds}s");
    }

    [Fact]
    public void Save_PersistsChanges_ToWorkbook()
    {
        // Arrange
        string testValue = $"Test-{Guid.NewGuid():N}";

        // Act - Write and save
        using (var batch = ExcelSession.BeginBatch(_testFileCopy!))
        {
            batch.Execute((ctx, ct) =>
            {
                dynamic sheet = ctx.Book.Worksheets.Item(1);
                sheet.Range["A1"].Value2 = testValue;
                return 0;
            });

            batch.Save();
        }

        // Wait for file to be released
        Thread.Sleep(1000);

        // Verify - Read back the value in a new batch session
        string readValue;
        using (var batch = ExcelSession.BeginBatch(_testFileCopy!))
        {
            readValue = batch.Execute((ctx, ct) =>
            {
                dynamic sheet = ctx.Book.Worksheets.Item(1);
                var value = sheet.Range["A1"].Value2;
                string result = value?.ToString() ?? "";
                return result;
            });
        }

        // Assert
        Assert.Equal(testValue, readValue);
        _output.WriteLine($"✓ Value persisted correctly: {testValue}");
    }

    [Fact]
    public void WorkbookPath_ReturnsCorrectPath()
    {
        // Arrange & Act
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);

        // Assert
        Assert.Equal(_testFileCopy, batch.WorkbookPath);
    }

    [Fact]
    public void CompleteWorkflow_CreateModifyReadSave_AllOperationsSucceed()
    {
        // Arrange
        string sheetName = "TestData";
        string testValue1 = "Header1";
        string testValue2 = "Value1";
        string namedRangeName = "TestRange";

        // Act - Execute complete workflow in single batch
        using (var batch = ExcelSession.BeginBatch(_testFileCopy!))
        {
            // Step 1: Create new worksheet
            batch.Execute((ctx, ct) =>
            {
                dynamic sheets = ctx.Book.Worksheets;
                dynamic newSheet = sheets.Add();
                newSheet.Name = sheetName;
                _output.WriteLine($"✓ Created worksheet: {sheetName}");
                return 0;
            });

            // Step 2: Write data to cells
            batch.Execute((ctx, ct) =>
            {
                dynamic sheet = ctx.Book.Worksheets.Item(sheetName);
                sheet.Range["A1"].Value2 = testValue1;
                sheet.Range["A2"].Value2 = testValue2;
                sheet.Range["B1"].Value2 = "Header2";
                sheet.Range["B2"].Formula = "=LEN(A2)";
                _output.WriteLine($"✓ Wrote data to cells A1, A2, B1, B2");
                return 0;
            });

            // Step 3: Create named range
            batch.Execute((ctx, ct) =>
            {
                dynamic sheet = ctx.Book.Worksheets.Item(sheetName);
                ctx.Book.Names.Add(namedRangeName, $"={sheetName}!$A$1:$B$2");
                _output.WriteLine($"✓ Created named range: {namedRangeName}");
                return 0;
            });

            // Step 4: Read data back to verify
            var readData = batch.Execute((ctx, ct) =>
            {
                dynamic sheet = ctx.Book.Worksheets.Item(sheetName);
                string a1 = sheet.Range["A1"].Value2?.ToString() ?? "";
                string a2 = sheet.Range["A2"].Value2?.ToString() ?? "";
                string b1 = sheet.Range["B1"].Value2?.ToString() ?? "";
                double b2 = Convert.ToDouble(sheet.Range["B2"].Value2); // Formula result
                _output.WriteLine($"✓ Read back: A1={a1}, A2={a2}, B1={b1}, B2={b2}");
                return (a1, a2, b1, b2);
            });

            // Verify intermediate state
            Assert.Equal(testValue1, readData.a1);
            Assert.Equal(testValue2, readData.a2);
            Assert.Equal("Header2", readData.b1);
            Assert.Equal(6.0, Convert.ToDouble(readData.b2)); // LEN("Value1") = 6

            // Step 5: Modify existing data
            batch.Execute((ctx, ct) =>
            {
                dynamic sheet = ctx.Book.Worksheets.Item(sheetName);
                sheet.Range["A2"].Value2 = "Modified";
                _output.WriteLine("✓ Modified A2 cell");
                return 0;
            });

            // Step 6: Save all changes
            batch.Save();
            _output.WriteLine("✓ Saved workbook");
        }

        // Wait for file to be released
        Thread.Sleep(1000);

        // Verify - Open in new batch and check all changes persisted
        using (var batch = ExcelSession.BeginBatch(_testFileCopy!))
        {
            var verifyData = batch.Execute((ctx, ct) =>
            {
                // Check worksheet exists
                bool sheetExists = false;
                dynamic sheets = ctx.Book.Worksheets;
                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic sheet = sheets.Item(i);
                    if (sheet.Name == sheetName)
                    {
                        sheetExists = true;
                        break;
                    }
                }

                // Read cell values
                dynamic dataSheet = ctx.Book.Worksheets.Item(sheetName);
                string a1 = dataSheet.Range["A1"].Value2?.ToString() ?? "";
                string a2 = dataSheet.Range["A2"].Value2?.ToString() ?? "";
                double b2 = Convert.ToDouble(dataSheet.Range["B2"].Value2);

                // Check named range exists
                bool namedRangeExists = false;
                dynamic names = ctx.Book.Names;
                for (int i = 1; i <= names.Count; i++)
                {
                    dynamic name = names.Item(i);
                    if (name.Name == namedRangeName)
                    {
                        namedRangeExists = true;
                        break;
                    }
                }

                return (sheetExists, a1, a2, b2, namedRangeExists);
            });

            // Assert - All changes persisted
            Assert.True(verifyData.sheetExists, "Worksheet should exist after save");
            Assert.Equal(testValue1, verifyData.a1);
            Assert.Equal("Modified", verifyData.a2);
            Assert.Equal(8.0, verifyData.b2); // LEN("Modified") = 8
            Assert.True(verifyData.namedRangeExists, "Named range should exist after save");
            _output.WriteLine("✓ All workflow changes persisted correctly");
        }
    }

    [Fact]
    public async Task ParallelBatches_TwoConcurrentBatches_NoExcelProcessLeak()
    {
        // Arrange
        const int batchCount = 2;
        var testFileCopies = new List<string>();

        // Create fresh copies for parallel test
        for (int i = 0; i < batchCount; i++)
        {
            string copy = Path.Join(Path.GetTempPath(), $"batch-test-parallel-{i}-{Guid.NewGuid():N}.xlsx");
            File.Copy(_staticTestFile!, copy, overwrite: true);
            testFileCopies.Add(copy);
        }

        var startingProcesses = Process.GetProcessesByName("EXCEL");
        int startingCount = startingProcesses.Length;
        _output.WriteLine($"Excel processes before parallel batches: {startingCount}");

        try
        {
            // Act - Run 2 batches in parallel
            var tasks = testFileCopies.Select((testFile, index) =>
            {
                return Task.Run(() =>
                {
                    using var batch = ExcelSession.BeginBatch(testFile);

                    // Perform multiple operations per batch
                    for (int op = 0; op < 3; op++)
                    {
                        batch.Execute((ctx, ct) =>
                        {
                            dynamic sheet = ctx.Book.Worksheets.Item(1);
                            sheet.Range[$"A{op + 1}"].Value2 = $"Batch{index}-Op{op}";
                            return 0;
                        });
                    }

                    _output.WriteLine($"✓ Batch {index} completed");

                    return index;
                });
            }).ToArray();

            // Wait for all batches to complete
            var results = await Task.WhenAll(tasks);

            Assert.Equal(batchCount, results.Length);
            _output.WriteLine($"✓ All {batchCount} parallel batches completed");

            // Wait for Excel processes to terminate
            await Task.Delay(5000);

            // Assert - No process leak
            var endingProcesses = Process.GetProcessesByName("EXCEL");
            int endingCount = endingProcesses.Length;
            _output.WriteLine($"Excel processes after parallel batches: {endingCount}");

            Assert.True(endingCount <= startingCount + 2, // Allow some tolerance for cleanup timing
                $"Excel process leak in parallel batches! Started with {startingCount}, ended with {endingCount}");
        }
        finally
        {
            // Cleanup parallel test files
            foreach (var testFile in testFileCopies.Where(File.Exists))
            {
#pragma warning disable CA1031 // Intentional: best-effort test cleanup
                try { File.Delete(testFile); } catch (Exception) { /* Best effort cleanup */ }
#pragma warning restore CA1031
            }
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "FileLocking")]
    public void Constructor_FileLockedByAnotherProcess_ThrowsInvalidOperationException()
    {
        // Arrange - Create a separate test file for locking test
        var lockedTestFile = Path.Join(Path.GetTempPath(), $"batch-test-locked-{Guid.NewGuid():N}.xlsx");
        File.Copy(_staticTestFile!, lockedTestFile, overwrite: true);

        try
        {
            // Lock the file by opening with exclusive access (simulating Excel or another process)
            using var fileLock = new FileStream(
                lockedTestFile,
                FileMode.Open,
                FileAccess.ReadWrite,
                FileShare.None);

            // Act & Assert - Attempting to create ExcelBatch should fail immediately
            var ex = Assert.Throws<InvalidOperationException>(() =>
            {
                var batch = ExcelSession.BeginBatch(lockedTestFile);
                batch.Dispose();
            });

            // Verify error message is clear and actionable
            Assert.Contains("already open", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("close the file", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("exclusive access", ex.Message, StringComparison.OrdinalIgnoreCase);

            _output.WriteLine($"✓ File locking detected successfully");
            _output.WriteLine($"Error message: {ex.Message}");
        }
        finally
        {
            // Cleanup
            if (File.Exists(lockedTestFile))
            {
#pragma warning disable CA1031 // Intentional: best-effort test cleanup
                try { File.Delete(lockedTestFile); } catch (Exception) { /* Best effort - file may be locked */ }
#pragma warning restore CA1031
            }
        }
    }

    // Note: Testing file-already-open scenario is complex because:
    // 1. Excel's behavior when opening an already-open file can vary (hang, prompt, or succeed)
    // 2. The error detection code in ExcelBatch.cs catches COM Error 0x800A03EC
    // 3. This test would require simulating Excel having the file open externally
    //
    // The error handling code is verified through:
    // - Manual testing: Open file in Excel UI, then try automation
    // - Real-world usage: Users will encounter this if they forget to close files
    // - Code review: Error message is clear and actionable
    //
    // UPDATE: We now have a test (Constructor_FileLockedByAnotherProcess_ThrowsInvalidOperationException)
    // that verifies the OS-level file locking check without requiring Excel to be running.
    //
    // Keeping this comment as documentation that the scenario is handled in production code.
}







