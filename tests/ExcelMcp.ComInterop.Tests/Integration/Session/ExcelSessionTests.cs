using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit.Session;

/// <summary>
/// Integration tests for ExcelSession - verifies STA threading and COM cleanup.
/// Tests that Excel.exe processes are properly terminated after operations.
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test ExcelSession.ExecuteAsync() COM cleanup
/// - ✅ Test ExcelSession.CreateNewAsync() COM cleanup
/// - ✅ Test ExcelSession.BeginBatchAsync() factory method
/// - ✅ Verify Excel.exe process termination (no leaks)
///
/// NOTE: ExcelSession handles all GC cleanup automatically in its cleanup code.
/// Tests only need to wait for async disposal and process termination timing.
///
/// IMPORTANT: These tests spawn and terminate Excel processes (side effects).
/// They run OnDemand only to avoid interference with normal test runs.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "ExcelSession")]
[Trait("RunType", "OnDemand")]
[Collection("Sequential")] // Disable parallelization to avoid COM interference
public class ExcelSessionTests
{
    private readonly ITestOutputHelper _output;

    public ExcelSessionTests(ITestOutputHelper output)
    {
        _output = output;
    }

    private async Task<string> CreateTempTestFileAsync()
    {
        string testFile = Path.Combine(Path.GetTempPath(), $"session-test-{Guid.NewGuid():N}.xlsx");
        await ExcelSession.CreateNewAsync(testFile, isMacroEnabled: false, (ctx, ct) =>
        {
            // File created, just return
            return ValueTask.FromResult(0);
        });
        return testFile;
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public async Task ExecuteAsync_SingleOperation_CleansUpExcelProcess()
    {
        // Arrange
        string testFile = await CreateTempTestFileAsync();
        var startingProcesses = Process.GetProcessesByName("EXCEL");
        int startingCount = startingProcesses.Length;

        _output.WriteLine($"Excel processes before: {startingCount}");

        try
        {
            // Act - Execute single operation
            await ExcelSession.ExecuteAsync(testFile, save: false, (ctx, ct) =>
            {
                dynamic sheet = ctx.Book.Worksheets.Item(1);
                var value = sheet.Range["A1"].Value2;
                _output.WriteLine($"Read A1: {value}");
                return ValueTask.FromResult(0);
            });

            // Wait for Excel process to fully terminate (ExecuteAsync handles GC cleanup)
            await Task.Delay(5000); // Increased delay for reliable process termination

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
    [Trait("RunType", "OnDemand")]
    public async Task CreateNewAsync_CreatesWorkbook_CleansUpExcelProcess()
    {
        // Arrange
        string tempFile = Path.Combine(Path.GetTempPath(), $"create-test-{Guid.NewGuid():N}.xlsx");
        var startingProcesses = Process.GetProcessesByName("EXCEL");
        int startingCount = startingProcesses.Length;

        _output.WriteLine($"Excel processes before: {startingCount}");

        try
        {
            // Act - Create new workbook
            await ExcelSession.CreateNewAsync(tempFile, isMacroEnabled: false, (ctx, ct) =>
            {
                dynamic sheet = ctx.Book.Worksheets.Item(1);
                sheet.Range["A1"].Value2 = "Test";
                return ValueTask.FromResult(0);
            });

            // Wait for Excel process to fully terminate (CreateNewAsync handles GC cleanup)
            await Task.Delay(2000);

            // Assert
            var endingProcesses = Process.GetProcessesByName("EXCEL");
            int endingCount = endingProcesses.Length;

            _output.WriteLine($"Excel processes after: {endingCount}");

            Assert.True(endingCount <= startingCount,
                $"Excel process leak in CreateNew! Started with {startingCount}, ended with {endingCount}");

            // Verify file was created
            Assert.True(File.Exists(tempFile), "Workbook file should exist");
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
    [Trait("RunType", "OnDemand")]
    public async Task BeginBatchAsync_ReturnsValidBatch_WithCorrectWorkbookPath()
    {
        // Arrange
        string testFile = await CreateTempTestFileAsync();

        try
        {
            // Act
            await using var batch = await ExcelSession.BeginBatchAsync(testFile);

            // Assert
            Assert.NotNull(batch);
            Assert.Equal(testFile, batch.WorkbookPath);
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }
}
