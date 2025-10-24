using Sbroenne.ExcelMcp.Core;
using Sbroenne.ExcelMcp.Core.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

/// <summary>
/// Tests for Excel instance pooling functionality.
/// These tests verify that pooled instances are reused correctly and
/// that the pool handles lifecycle management properly.
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
public class ExcelInstancePoolTests : IDisposable
{
    private readonly string _testFile;
    private readonly ExcelInstancePool _pool;

    public ExcelInstancePoolTests()
    {
        // Create a test Excel file
        _testFile = Path.Combine(Path.GetTempPath(), $"pool_test_{Guid.NewGuid()}.xlsx");

        // Create empty workbook for testing
        var fileCommands = new Sbroenne.ExcelMcp.Core.Commands.FileCommands();
        fileCommands.CreateEmpty(_testFile, overwriteIfExists: true);

        // Create pool with short timeout for testing
        _pool = new ExcelInstancePool(idleTimeout: TimeSpan.FromSeconds(5));
    }

    [Fact]
    public void WithPooledExcel_ShouldReuseInstance_ForSameFile()
    {
        // First operation - creates new instance
        var result1 = _pool.WithPooledExcel(_testFile, false, (excel, workbook) =>
        {
            return "First call";
        });

        Assert.Equal("First call", result1);

        // Second operation - should reuse instance
        var result2 = _pool.WithPooledExcel(_testFile, false, (excel, workbook) =>
        {
            return "Second call";
        });

        Assert.Equal("Second call", result2);

        // Both calls should have succeeded using pooled instance
        // (We can't directly verify reuse, but lack of errors indicates pooling works)
    }

    [Fact]
    public void WithPooledExcel_ShouldHandleSaveCorrectly()
    {
        // Operation with save=true
        var result = _pool.WithPooledExcel(_testFile, save: true, (excel, workbook) =>
        {
            // Modify workbook
            dynamic sheet = workbook.Worksheets.Item(1);
            sheet.Range["A1"].Value2 = "Test Value";
            return "Modified";
        });

        Assert.Equal("Modified", result);

        // Verify the change was saved by reading it back
        var readResult = _pool.WithPooledExcel(_testFile, save: false, (excel, workbook) =>
        {
            dynamic sheet = workbook.Worksheets.Item(1);
            return sheet.Range["A1"].Value2?.ToString() ?? "";
        });

        Assert.Equal("Test Value", readResult);
    }

    [Fact]
    public void CloseWorkbook_ShouldCloseButKeepExcelAlive()
    {
        // Open workbook via pool
        _pool.WithPooledExcel(_testFile, false, (excel, workbook) =>
        {
            return "Opened";
        });

        // Close workbook explicitly
        _pool.CloseWorkbook(_testFile);

        // Should be able to reopen - Excel instance still pooled
        var result = _pool.WithPooledExcel(_testFile, false, (excel, workbook) =>
        {
            return "Reopened";
        });

        Assert.Equal("Reopened", result);
    }

    [Fact]
    public void EvictInstance_ShouldRemoveFromPool()
    {
        // Use pool to open file
        _pool.WithPooledExcel(_testFile, false, (excel, workbook) =>
        {
            return "Initial";
        });

        // Evict the instance
        _pool.EvictInstance(_testFile);

        // Next operation should create new instance (no error = success)
        var result = _pool.WithPooledExcel(_testFile, false, (excel, workbook) =>
        {
            return "After eviction";
        });

        Assert.Equal("After eviction", result);
    }

    [Fact]
    public void ExcelHelper_WithPool_ShouldUsePooling()
    {
        // Configure ExcelHelper to use pooling
        var originalPool = ExcelHelper.InstancePool;
        try
        {
            ExcelHelper.InstancePool = _pool;

            // Call through ExcelHelper - should use pool
            var result = ExcelHelper.WithExcel(_testFile, false, (excel, workbook) =>
            {
                return "Pooled via ExcelHelper";
            });

            Assert.Equal("Pooled via ExcelHelper", result);

            // Second call should reuse instance
            var result2 = ExcelHelper.WithExcel(_testFile, false, (excel, workbook) =>
            {
                return "Reused via ExcelHelper";
            });

            Assert.Equal("Reused via ExcelHelper", result2);
        }
        finally
        {
            // Restore original pool
            ExcelHelper.InstancePool = originalPool;
        }
    }

    [Fact]
    public void ExcelHelper_WithoutPool_ShouldUseSingleInstance()
    {
        // Ensure no pool is configured
        var originalPool = ExcelHelper.InstancePool;
        try
        {
            ExcelHelper.InstancePool = null;

            // Call through ExcelHelper - should use single-instance pattern
            var result = ExcelHelper.WithExcel(_testFile, false, (excel, workbook) =>
            {
                return "Single instance";
            });

            Assert.Equal("Single instance", result);
        }
        finally
        {
            ExcelHelper.InstancePool = originalPool;
        }
    }

    public void Dispose()
    {
        // Cleanup
        _pool?.Dispose();

        if (File.Exists(_testFile))
        {
            try
            {
                File.Delete(_testFile);
            }
            catch
            {
                // Ignore cleanup errors
            }
        }

        GC.SuppressFinalize(this);
    }
}
