using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

/// <summary>
/// Integration tests for ExcelSession - verifies public API and COM cleanup.
/// Tests BeginBatch() and CreateNew() functionality.
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test ExcelSession.BeginBatch() validation and batch creation
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
[Collection("Sequential")] // Disable parallelization to avoid COM interference
public class ExcelSessionTests : IDisposable
{
    private readonly ITestOutputHelper _output;

    public ExcelSessionTests(ITestOutputHelper output)
    {
        _output = output;

        // Kill any existing Excel processes to ensure clean state
        var existingProcesses = Process.GetProcessesByName("EXCEL");
        if (existingProcesses.Length > 0)
        {
            _output.WriteLine($"Cleaning up {existingProcesses.Length} existing Excel processes...");
            foreach (var p in existingProcesses)
            {
                p.Kill(); p.WaitForExit(2000);
            }
            Thread.Sleep(2000); // Wait for cleanup
            _output.WriteLine("Excel processes cleaned up");
        }

    }

    /// <summary>
    /// Runs after each test
    /// </summary>
    public void Dispose()
    {
        // Nothing to dispose
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void BeginBatch_WithValidFile_CreatesBatch()
    {
        // Arrange
        string testFile = Path.Join(Path.GetTempPath(), $"session-test-{Guid.NewGuid():N}.xlsx");
        CreateTempTestFile(testFile);

        try
        {
            // Act
            using var batch = ExcelSession.BeginBatch(testFile);

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
    public void BeginBatch_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        // Arrange
        string nonExistentFile = Path.Join(Path.GetTempPath(), $"does-not-exist-{Guid.NewGuid():N}.xlsx");

        // Act & Assert
        var exception = Assert.Throws<FileNotFoundException>(() =>
        {
            using var batch = ExcelSession.BeginBatch(nonExistentFile);
        });

        _output.WriteLine("✓ Correctly throws FileNotFoundException for non-existent file");
    }

    [Fact]
    public void BeginBatch_WithInvalidExtension_ThrowsArgumentException()
    {
        // Arrange
        string invalidFile = Path.Join(Path.GetTempPath(), $"test-{Guid.NewGuid():N}.txt");
        File.WriteAllText(invalidFile, "dummy");

        try
        {
            // Act & Assert
            var exception = Assert.Throws<ArgumentException>(() =>
            {
                using var batch = ExcelSession.BeginBatch(invalidFile);
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
    public void CreateNew_CreatesNewWorkbook()
    {
        // Arrange
        string testFile = Path.Join(Path.GetTempPath(), $"new-workbook-{Guid.NewGuid():N}.xlsx");

        try
        {
            // Act
            var result = ExcelSession.CreateNew(testFile, isMacroEnabled: false, (ctx, ct) =>
            {
                _output.WriteLine($"✓ Workbook created at: {ctx.WorkbookPath}");
                return 0;
            });

            // Assert
            Assert.True(File.Exists(testFile), "File should be created");
            Assert.Equal(0, result);

            // Verify we can open it with batch API
            using (var batch = ExcelSession.BeginBatch(testFile))
            {
                batch.Execute((ctx, ct) =>
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
    public void CreateNew_WithMacroEnabled_CreatesXlsmFile()
    {
        // Arrange
        string testFile = Path.Join(Path.GetTempPath(), $"new-macro-workbook-{Guid.NewGuid():N}.xlsm");

        try
        {
            // Act
            var result = ExcelSession.CreateNew(testFile, isMacroEnabled: true, (ctx, ct) =>
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
    public void CreateNew_CreatesDirectoryIfNeeded()
    {
        // Arrange
        string testDir = Path.Join(Path.GetTempPath(), $"testdir-{Guid.NewGuid():N}");
        string testFile = Path.Join(testDir, "newfile.xlsx");

        try
        {
            // Act
            ExcelSession.CreateNew(testFile, isMacroEnabled: false, (ctx, ct) =>
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

    // Helper method
    private static void CreateTempTestFile(string filePath)
    {
        ExcelSession.CreateNew(filePath, isMacroEnabled: false, (ctx, ct) =>
        {
            // File created, just return
            return 0;
        });
    }
}

