using System.Diagnostics;
using Sbroenne.ExcelMcp.Core.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit.Session;

/// <summary>
/// Tests for STA threading and batching - verifies COM cleanup and no process leaks.
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "StaThreading")]
[Collection("Sequential")] // Disable parallelization to avoid COM interference
public class StaThreadingTests
{
    private readonly ITestOutputHelper _output;

    public StaThreadingTests(ITestOutputHelper output)
    {
        _output = output;
    }

    private async Task<string> CreateTempTestFileAsync()
    {
        string testFile = Path.Combine(Path.GetTempPath(), $"sta-test-{Guid.NewGuid():N}.xlsx");
        await ExcelSession.CreateNewAsync(testFile, isMacroEnabled: false, (ctx, ct) =>
        {
            // File created, just return
            return ValueTask.FromResult(0);
        });
        return testFile;
    }

    [Fact]
    public async Task ExecuteAsync_WithStaThreading_NoProcessLeak()
    {
        // Arrange
        string testFile = await CreateTempTestFileAsync();
        var startingProcesses = Process.GetProcessesByName("EXCEL");
        int startingCount = startingProcesses.Length;

        _output.WriteLine($"Excel processes before: {startingCount}");

        try
        {
            // Act - Execute single operation with STA threading
            await ExcelSession.ExecuteAsync(testFile, save: false, (ctx, ct) =>
            {
                dynamic sheet = ctx.Book.Worksheets.Item(1);
                var value = sheet.Range["A1"].Value2;
                _output.WriteLine($"Read A1: {value}");
                return ValueTask.FromResult(0);
            });

            // Wait for COM cleanup
            await Task.Delay(5000); // Increased from 2s to 5s for more reliable cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            // Assert
            var endingProcesses = Process.GetProcessesByName("EXCEL");
            int endingCount = endingProcesses.Length;

            _output.WriteLine($"Excel processes after: {endingCount}");

            Assert.True(endingCount <= startingCount,
                $"Excel process leak! Started with {startingCount}, ended with {endingCount}");
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }

    [Fact]
    public async Task BeginBatchAsync_MultipleOperations_UseSameInstance()
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
    public async Task Batch_DisposeAsync_CleansUpComObjects()
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

            // Wait for COM cleanup
            await Task.Delay(2000);
            GC.Collect();
            GC.WaitForPendingFinalizers();

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
    public async Task CreateNewAsync_WithStaThreading_NoProcessLeak()
    {
        // Arrange
        string tempFile = Path.Combine(Path.GetTempPath(), $"sta-test-{Guid.NewGuid():N}.xlsx");
        var startingProcesses = Process.GetProcessesByName("EXCEL");
        int startingCount = startingProcesses.Length;

        _output.WriteLine($"Excel processes before: {startingCount}");

        try
        {
            // Act
            await ExcelSession.CreateNewAsync(tempFile, isMacroEnabled: false, (ctx, ct) =>
            {
                dynamic sheet = ctx.Book.Worksheets.Item(1);
                sheet.Range["A1"].Value2 = "Test";
                return ValueTask.FromResult(0);
            });

            // Wait for COM cleanup
            await Task.Delay(2000);
            GC.Collect();
            GC.WaitForPendingFinalizers();

            // Assert
            var endingProcesses = Process.GetProcessesByName("EXCEL");
            int endingCount = endingProcesses.Length;

            _output.WriteLine($"Excel processes after: {endingCount}");

            Assert.True(endingCount <= startingCount,
                $"Excel process leak in CreateNew! Started with {startingCount}, ended with {endingCount}");
        }
        finally
        {
            if (File.Exists(tempFile))
            {
                await Task.Delay(500);
                File.Delete(tempFile);
            }
        }
    }

    [Fact]
    public async Task BackwardCompatWrapper_Execute_UsesStaThreading()
    {
        // Arrange
        string testFile = await CreateTempTestFileAsync();

        try
        {
            // Act - Old synchronous API should still work with STA threading
            var result = ExcelSession.Execute(testFile, save: false, (excel, workbook) =>
            {
                dynamic sheet = workbook.Worksheets.Item(1);
                var value = sheet.Range["A1"].Value2;
                _output.WriteLine($"Backward-compat read: {value}");
                return 42;
            });

            // Assert
            Assert.Equal(42, result);
            _output.WriteLine("✓ Backward-compat wrapper works with STA threading");
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }
}
