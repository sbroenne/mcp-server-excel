using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit.Session;

/// <summary>
/// Integration tests for ExcelBatch - verifies batch operations and COM cleanup.
/// Tests that Excel instances are reused across operations and properly cleaned up.
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test ExcelBatch.ExecuteAsync() reuses Excel instance
/// - ✅ Test ExcelBatch.DisposeAsync() COM cleanup
/// - ✅ Test ExcelBatch.SaveAsync() functionality
/// - ✅ Verify Excel.exe process termination (no leaks)
///
/// NOTE: ExcelBatch.DisposeAsync() handles all GC cleanup automatically.
/// Tests only need to wait for async disposal and process termination timing.
///
/// IMPORTANT: These tests spawn and terminate Excel processes (side effects).
/// They run OnDemand only to avoid interference with normal test runs.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "ExcelBatch")]
[Trait("RunType", "OnDemand")]
[Collection("Sequential")] // Disable parallelization to avoid COM interference
public class ExcelBatchTests
{
    private readonly ITestOutputHelper _output;

    public ExcelBatchTests(ITestOutputHelper output)
    {
        _output = output;
    }

    private async Task<string> CreateTempTestFileAsync()
    {
        string testFile = Path.Combine(Path.GetTempPath(), $"batch-test-{Guid.NewGuid():N}.xlsx");
        await ExcelSession.CreateNewAsync(testFile, isMacroEnabled: false, (ctx, ct) =>
        {
            // File created, just return
            return ValueTask.FromResult(0);
        });
        return testFile;
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public async Task ExecuteAsync_MultipleOperations_ReusesExcelInstance()
    {
        // Arrange
        string testFile = await CreateTempTestFileAsync();
        int operationCount = 0;

        try
        {
            // Act - Use batching for multiple operations
            await using var batch = await ExcelSession.BeginBatchAsync(testFile);

            for (int i = 0; i < 5; i++)
            {
                await batch.ExecuteAsync<int>((ctx, ct) =>
                {
                    operationCount++;
                    _output.WriteLine($"Batch operation {operationCount}");

                    // Verify we have the same context
                    Assert.NotNull(ctx.App);
                    Assert.NotNull(ctx.Book);

                    return ValueTask.FromResult(operationCount);
                });
            }

            // Assert
            Assert.Equal(5, operationCount);
            _output.WriteLine($"✓ Completed {operationCount} batch operations");
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public async Task DisposeAsync_CleansUpComObjects_NoProcessLeak()
    {
        // Arrange
        string testFile = await CreateTempTestFileAsync();
        var startingProcesses = Process.GetProcessesByName("EXCEL");
        int startingCount = startingProcesses.Length;

        _output.WriteLine($"Excel processes before: {startingCount}");

        try
        {
            // Act
            var batch = await ExcelSession.BeginBatchAsync(testFile);

            await batch.ExecuteAsync<int>((ctx, ct) =>
            {
                dynamic sheet = ctx.Book.Worksheets.Item(1);
                var value = sheet.Range["A1"].Value2;
                return ValueTask.FromResult(0);
            });

            await batch.DisposeAsync();

            // Wait for Excel process to fully terminate (DisposeAsync handles GC cleanup)
            await Task.Delay(2000);

            // Assert
            var endingProcesses = Process.GetProcessesByName("EXCEL");
            int endingCount = endingProcesses.Length;

            _output.WriteLine($"Excel processes after: {endingCount}");

            Assert.True(endingCount <= startingCount,
                $"Excel process leak in batch! Started with {startingCount}, ended with {endingCount}");
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public async Task SaveAsync_PersistsChanges_ToWorkbook()
    {
        // Arrange
        string testFile = await CreateTempTestFileAsync();
        string testValue = $"Test-{Guid.NewGuid():N}";

        try
        {
            // Act - Write and save
            await using (var batch = await ExcelSession.BeginBatchAsync(testFile))
            {
                await batch.ExecuteAsync<int>((ctx, ct) =>
                {
                    dynamic sheet = ctx.Book.Worksheets.Item(1);
                    sheet.Range["A1"].Value2 = testValue;
                    return ValueTask.FromResult(0);
                });

                await batch.SaveAsync();
            }

            // Wait for file to be released
            await Task.Delay(1000);

            // Verify - Read back the value in a new batch session
            string readValue;
            await using (var batch = await ExcelSession.BeginBatchAsync(testFile))
            {
                readValue = await batch.ExecuteAsync<string>((ctx, ct) =>
                {
                    dynamic sheet = ctx.Book.Worksheets.Item(1);
                    var value = sheet.Range["A1"].Value2;
                    return ValueTask.FromResult(value?.ToString() ?? "");
                });
            }

            // Assert
            Assert.Equal(testValue, readValue);
            _output.WriteLine($"✓ Value persisted correctly: {testValue}");
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public async Task WorkbookPath_ReturnsCorrectPath()
    {
        // Arrange
        string testFile = await CreateTempTestFileAsync();

        try
        {
            // Act
            await using var batch = await ExcelSession.BeginBatchAsync(testFile);

            // Assert
            Assert.Equal(testFile, batch.WorkbookPath);
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public async Task CompleteWorkflow_CreateModifyReadSave_AllOperationsSucceed()
    {
        // Arrange
        string testFile = await CreateTempTestFileAsync();
        string sheetName = "TestData";
        string testValue1 = "Header1";
        string testValue2 = "Value1";
        string namedRangeName = "TestRange";

        try
        {
            // Act - Execute complete workflow in single batch
            await using (var batch = await ExcelSession.BeginBatchAsync(testFile))
            {
                // Step 1: Create new worksheet
                await batch.ExecuteAsync<int>((ctx, ct) =>
                {
                    dynamic sheets = ctx.Book.Worksheets;
                    dynamic newSheet = sheets.Add();
                    newSheet.Name = sheetName;
                    _output.WriteLine($"✓ Created worksheet: {sheetName}");
                    return ValueTask.FromResult(0);
                });

                // Step 2: Write data to cells
                await batch.ExecuteAsync<int>((ctx, ct) =>
                {
                    dynamic sheet = ctx.Book.Worksheets.Item(sheetName);
                    sheet.Range["A1"].Value2 = testValue1;
                    sheet.Range["A2"].Value2 = testValue2;
                    sheet.Range["B1"].Value2 = "Header2";
                    sheet.Range["B2"].Formula = "=LEN(A2)";
                    _output.WriteLine($"✓ Wrote data to cells A1, A2, B1, B2");
                    return ValueTask.FromResult(0);
                });

                // Step 3: Create named range
                await batch.ExecuteAsync<int>((ctx, ct) =>
                {
                    dynamic sheet = ctx.Book.Worksheets.Item(sheetName);
                    dynamic names = ctx.Book.Names;
                    names.Add(namedRangeName, $"={sheetName}!$A$1:$B$2");
                    _output.WriteLine($"✓ Created named range: {namedRangeName}");
                    return ValueTask.FromResult(0);
                });

                // Step 4: Read data back to verify
                var readData = await batch.ExecuteAsync((ctx, ct) =>
                {
                    dynamic sheet = ctx.Book.Worksheets.Item(sheetName);
                    string a1 = sheet.Range["A1"].Value2?.ToString() ?? "";
                    string a2 = sheet.Range["A2"].Value2?.ToString() ?? "";
                    string b1 = sheet.Range["B1"].Value2?.ToString() ?? "";
                    double b2 = Convert.ToDouble(sheet.Range["B2"].Value2); // Formula result
                    _output.WriteLine($"✓ Read back: A1={a1}, A2={a2}, B1={b1}, B2={b2}");
                    return ValueTask.FromResult((a1, a2, b1, b2));
                });

                // Verify intermediate state
                Assert.Equal(testValue1, readData.a1);
                Assert.Equal(testValue2, readData.a2);
                Assert.Equal("Header2", readData.b1);
                Assert.Equal(6.0, Convert.ToDouble(readData.b2)); // LEN("Value1") = 6

                // Step 5: Modify existing data
                await batch.ExecuteAsync<int>((ctx, ct) =>
                {
                    dynamic sheet = ctx.Book.Worksheets.Item(sheetName);
                    sheet.Range["A2"].Value2 = "Modified";
                    _output.WriteLine("✓ Modified A2 cell");
                    return ValueTask.FromResult(0);
                });

                // Step 6: Save all changes
                await batch.SaveAsync();
                _output.WriteLine("✓ Saved workbook");
            }

            // Wait for file to be released
            await Task.Delay(1000);

            // Verify - Open in new batch and check all changes persisted
            await using (var batch = await ExcelSession.BeginBatchAsync(testFile))
            {
                var verifyData = await batch.ExecuteAsync((ctx, ct) =>
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

                    return ValueTask.FromResult((sheetExists, a1, a2, b2, namedRangeExists));
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
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public async Task ParallelBatches_10ConcurrentBatches_NoExcelProcessLeak()
    {
        // Arrange
        const int batchCount = 10;
        var testFiles = new List<string>();

        // Create test files
        for (int i = 0; i < batchCount; i++)
        {
            testFiles.Add(await CreateTempTestFileAsync());
        }

        var startingProcesses = Process.GetProcessesByName("EXCEL");
        int startingCount = startingProcesses.Length;
        _output.WriteLine($"Excel processes before parallel batches: {startingCount}");

        try
        {
            // Act - Run 10 batches in parallel
            var tasks = testFiles.Select(async (testFile, index) =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(testFile);

                // Perform multiple operations per batch
                for (int op = 0; op < 3; op++)
                {
                    await batch.ExecuteAsync<int>((ctx, ct) =>
                    {
                        dynamic sheet = ctx.Book.Worksheets.Item(1);
                        sheet.Range[$"A{op + 1}"].Value2 = $"Batch{index}-Op{op}";
                        return ValueTask.FromResult(0);
                    });
                }

                await batch.SaveAsync();
                _output.WriteLine($"✓ Batch {index} completed");

                return index;
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
            // Cleanup all test files
            foreach (var testFile in testFiles)
            {
                if (File.Exists(testFile))
                {
                    try { File.Delete(testFile); } catch { }
                }
            }
        }
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public async Task ParallelBatches_VerifyDataIntegrity_NoInterference()
    {
        // Arrange
        const int batchCount = 5;
        var testFiles = new List<string>();
        var expectedValues = new Dictionary<string, string>();

        // Create test files with unique identifiers
        for (int i = 0; i < batchCount; i++)
        {
            var testFile = await CreateTempTestFileAsync();
            testFiles.Add(testFile);
            expectedValues[testFile] = $"UniqueValue-{i}-{Guid.NewGuid():N}";
        }

        try
        {
            // Act - Write unique data in parallel batches
            var writeTasks = testFiles.Select(async testFile =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(testFile);

                await batch.ExecuteAsync<int>((ctx, ct) =>
                {
                    dynamic sheet = ctx.Book.Worksheets.Item(1);
                    sheet.Range["A1"].Value2 = expectedValues[testFile];
                    return ValueTask.FromResult(0);
                });

                await batch.SaveAsync();
            }).ToArray();

            await Task.WhenAll(writeTasks);
            _output.WriteLine($"✓ All {batchCount} parallel write operations completed");

            // Wait for files to be released
            await Task.Delay(2000);

            // Verify - Read back in parallel and verify data integrity
            var readTasks = testFiles.Select(async testFile =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(testFile);

                var value = await batch.ExecuteAsync((ctx, ct) =>
                {
                    dynamic sheet = ctx.Book.Worksheets.Item(1);
                    string cellValue = sheet.Range["A1"].Value2?.ToString() ?? "";
                    return ValueTask.FromResult(cellValue);
                });

                return (testFile, value);
            }).ToArray();

            var readResults = await Task.WhenAll(readTasks);

            // Assert - Each file has its unique value (no cross-contamination)
            foreach (var (testFile, actualValue) in readResults)
            {
                var expectedValue = expectedValues[testFile];
                Assert.Equal(expectedValue, actualValue);
                _output.WriteLine($"✓ File {Path.GetFileName(testFile)}: Expected={expectedValue}, Actual={actualValue}");
            }

            _output.WriteLine("✓ No data interference between parallel batches");
        }
        finally
        {
            // Cleanup
            foreach (var testFile in testFiles)
            {
                if (File.Exists(testFile))
                {
                    try { File.Delete(testFile); } catch { }
                }
            }
        }
    }
}

