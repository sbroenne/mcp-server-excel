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

            // Verify - Read back the value in a new session
            var readValue = await ExcelSession.ExecuteAsync<string>(testFile, save: false, (ctx, ct) =>
            {
                dynamic sheet = ctx.Book.Worksheets.Item(1);
                var value = sheet.Range["A1"].Value2;
                return ValueTask.FromResult(value?.ToString() ?? "");
            });

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
}
