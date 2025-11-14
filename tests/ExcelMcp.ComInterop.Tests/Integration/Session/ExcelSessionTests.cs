using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

/// <summary>
/// Integration tests for ExcelSession - verifies public API and COM cleanup.
/// Tests BeginBatchAsync() and CreateNewAsync() functionality.
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test ExcelSession.BeginBatchAsync() validation and batch creation
/// - ✅ Test ExcelSession.CreateNew() file creation
/// - ✅ Verify Excel.exe process termination (no leaks)
///
/// NOTE: ExcelSession methods handle all GC cleanup automatically.
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
public class ExcelSessionTests : IAsyncLifetime
{
    private readonly ITestOutputHelper _output;

    public ExcelSessionTests(ITestOutputHelper output)
    {
        _output = output;
    }

    /// <summary>
    /// Runs before each test to ensure clean Excel process state
    /// </summary>
    public async Task InitializeAsync()
    {
        // Kill any existing Excel processes to ensure clean state
        try
        {
            var existingProcesses = Process.GetProcessesByName("EXCEL");
            if (existingProcesses.Length > 0)
            {
                _output.WriteLine($"Cleaning up {existingProcesses.Length} existing Excel processes...");
                foreach (var p in existingProcesses)
                {
                    try { p.Kill(); p.WaitForExit(2000); } catch { }
                }
                await Task.Delay(2000); // Wait for cleanup
                _output.WriteLine("Excel processes cleaned up");
            }
        }
        catch { }
    }

    /// <summary>
    /// Runs after each test
    /// </summary>
    public Task DisposeAsync()
    {
        return Task.CompletedTask;
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public async Task BeginBatchAsync_WithValidFile_CreatesBatch()
    {
        // Arrange
        string testFile = Path.Join(Path.GetTempPath(), $"session-test-{Guid.NewGuid():N}.xlsx");
        await CreateTempTestFileAsync(testFile);

        try
        {
            // Act
            await using var batch = await ExcelSession.BeginBatchAsync(testFile);

            // Assert
            Assert.NotNull(batch);
            Assert.Equal(testFile, batch.WorkbookPath);

            _output.WriteLine($"✓ Batch created successfully for: {Path.GetFileName(testFile)}");
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public async Task BeginBatchAsync_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        // Arrange
        string nonExistentFile = Path.Join(Path.GetTempPath(), $"does-not-exist-{Guid.NewGuid():N}.xlsx");

        // Act & Assert
        await Assert.ThrowsAsync<FileNotFoundException>(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(nonExistentFile);
        });

        _output.WriteLine("✓ Correctly throws FileNotFoundException for non-existent file");
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public async Task BeginBatchAsync_WithInvalidExtension_ThrowsArgumentException()
    {
        // Arrange
        string invalidFile = Path.Join(Path.GetTempPath(), $"test-{Guid.NewGuid():N}.txt");
        File.WriteAllText(invalidFile, "dummy");

        try
        {
            // Act & Assert
            var exception = await Assert.ThrowsAsync<ArgumentException>(async () =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(invalidFile);
            });

            Assert.Contains("Invalid file extension", exception.Message);
            _output.WriteLine("✓ Correctly rejects non-Excel file extension");
        }
        finally
        {
            if (File.Exists(invalidFile)) File.Delete(invalidFile);
        }
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public async Task BeginBatchAsync_DisposesCorrectly_NoExcelProcessLeak()
    {
        // Arrange
        string testFile = Path.Join(Path.GetTempPath(), $"session-test-{Guid.NewGuid():N}.xlsx");
        await CreateTempTestFileAsync(testFile);

        var startingProcesses = Process.GetProcessesByName("EXCEL");
        int startingCount = startingProcesses.Length;
        _output.WriteLine($"Excel processes before: {startingCount}");

        try
        {
            // Act - Create and dispose batch
            await using (var batch = await ExcelSession.BeginBatchAsync(testFile))
            {
                await batch.Execute((ctx, ct) =>
                {
                    dynamic sheet = ctx.Book.Worksheets.Item(1);
                    var value = sheet.Range["A1"].Value2;
                    _output.WriteLine($"Read A1: {value}");
                    return 0;
                });
            }

            // Wait for Excel process to fully terminate
            await Task.Delay(5000);

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
    public async Task CreateNewAsync_CreatesNewWorkbook()
    {
        // Arrange
        string testFile = Path.Join(Path.GetTempPath(), $"new-workbook-{Guid.NewGuid():N}.xlsx");

        try
        {
            // Act
            var result = await ExcelSession.CreateNew(testFile, isMacroEnabled: false, (ctx, ct) =>
            {
                _output.WriteLine($"✓ Workbook created at: {ctx.WorkbookPath}");
                return 0;
            });

            // Assert
            Assert.True(File.Exists(testFile), "File should be created");
            Assert.Equal(0, result);

            // Verify we can open it with batch API
            await using (var batch = await ExcelSession.BeginBatchAsync(testFile))
            {
                await batch.Execute((ctx, ct) =>
                {
                    Assert.NotNull(ctx.Book);
                    _output.WriteLine("✓ Can open created workbook with batch API");
                    return 0;
                });
            }
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public async Task CreateNewAsync_WithMacroEnabled_CreatesXlsmFile()
    {
        // Arrange
        string testFile = Path.Join(Path.GetTempPath(), $"new-macro-workbook-{Guid.NewGuid():N}.xlsm");

        try
        {
            // Act
            var result = await ExcelSession.CreateNew(testFile, isMacroEnabled: true, (ctx, ct) =>
            {
                _output.WriteLine($"✓ Macro-enabled workbook created at: {ctx.WorkbookPath}");
                return 0;
            });

            // Assert
            Assert.True(File.Exists(testFile), "XLSM file should be created");
            Assert.Equal(".xlsm", Path.GetExtension(testFile).ToLowerInvariant());
            _output.WriteLine("✓ Correctly created .xlsm file");
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public async Task CreateNewAsync_CreatesDirectoryIfNeeded()
    {
        // Arrange
        string testDir = Path.Join(Path.GetTempPath(), $"testdir-{Guid.NewGuid():N}");
        string testFile = Path.Join(testDir, "newfile.xlsx");

        try
        {
            // Act
            await ExcelSession.CreateNew(testFile, isMacroEnabled: false, (ctx, ct) =>
            {
                return 0;
            });

            // Assert
            Assert.True(Directory.Exists(testDir), "Directory should be created");
            Assert.True(File.Exists(testFile), "File should be created in new directory");
            _output.WriteLine("✓ Correctly created directory and file");
        }
        finally
        {
            if (Directory.Exists(testDir)) Directory.Delete(testDir, recursive: true);
        }
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public async Task CreateNewAsync_NoExcelProcessLeak()
    {
        // Arrange
        string testFile = Path.Join(Path.GetTempPath(), $"new-workbook-{Guid.NewGuid():N}.xlsx");

        var startingProcesses = Process.GetProcessesByName("EXCEL");
        int startingCount = startingProcesses.Length;
        _output.WriteLine($"Excel processes before: {startingCount}");

        try
        {
            // Act
            await ExcelSession.CreateNew(testFile, isMacroEnabled: false, (ctx, ct) =>
            {
                return 0;
            });

            // Force garbage collection to help COM cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            // Wait for Excel process to fully terminate
            await Task.Delay(7000); // Increased from 5000

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

    // Helper method
    private static async Task CreateTempTestFileAsync(string filePath)
    {
        await ExcelSession.CreateNew(filePath, isMacroEnabled: false, (ctx, ct) =>
        {
            // File created, just return
            return 0;
        });
    }
}
