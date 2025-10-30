using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests for Sheet lifecycle operations.
/// These tests require Excel installation and validate Core worksheet lifecycle management.
/// Tests use Core commands directly (not through CLI wrapper).
/// Data operations (read, write, clear) moved to RangeCommandsTests.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Worksheets")]
public class SheetCommandsTests : IDisposable
{
    private readonly ISheetCommands _sheetCommands;
    private readonly IFileCommands _fileCommands;
    private readonly string _testExcelFile;
    private readonly string _tempDir;
    private bool _disposed;

    public SheetCommandsTests()
    {
        _sheetCommands = new SheetCommands();
        _fileCommands = new FileCommands();

        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_Sheet_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _testExcelFile = Path.Combine(_tempDir, "TestWorkbook.xlsx");

        // Create test Excel file
        CreateTestExcelFile();
    }

    private void CreateTestExcelFile()
    {
        var result = _fileCommands.CreateEmptyAsync(_testExcelFile, overwriteIfExists: false).GetAwaiter().GetResult();
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}");
        }
    }

    [Fact]
    public async Task List_WithValidFile_ReturnsSuccessResult()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _sheetCommands.ListAsync(batch);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Worksheets);
        Assert.NotEmpty(result.Worksheets); // New Excel file has Sheet1
    }

    [Fact]
    public async Task Create_WithValidName_ReturnsSuccessResult()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _sheetCommands.CreateAsync(batch, "TestSheet");
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");

        // Verify sheet actually exists
        var listResult = await _sheetCommands.ListAsync(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Worksheets, w => w.Name == "TestSheet");
    }

    [Fact]
    public async Task List_AfterCreate_ShowsNewSheet()
    {
        // Arrange - Create sheet
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            await _sheetCommands.CreateAsync(batch, "TestSheet");
            await batch.SaveAsync();
        }

        // Act - List sheets
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _sheetCommands.ListAsync(batch);

            // Assert
            Assert.True(result.Success);
            Assert.Contains(result.Worksheets, w => w.Name == "TestSheet");
        }
    }

    [Fact]
    public async Task Rename_WithValidNames_ReturnsSuccessResult()
    {
        // Arrange
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            await _sheetCommands.CreateAsync(batch, "OldName");
            await batch.SaveAsync();
        }

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _sheetCommands.RenameAsync(batch, "OldName", "NewName");
            await batch.SaveAsync();

            // Assert
            Assert.True(result.Success);

            // Verify rename actually happened
            var listResult = await _sheetCommands.ListAsync(batch);
            Assert.True(listResult.Success);
            Assert.DoesNotContain(listResult.Worksheets, w => w.Name == "OldName");
            Assert.Contains(listResult.Worksheets, w => w.Name == "NewName");
        }
    }

    [Fact]
    public async Task Delete_WithExistingSheet_ReturnsSuccessResult()
    {
        // Arrange
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            await _sheetCommands.CreateAsync(batch, "ToDelete");
            await batch.SaveAsync();
        }

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _sheetCommands.DeleteAsync(batch, "ToDelete");
            await batch.SaveAsync();

            // Assert
            Assert.True(result.Success);

            // Verify sheet is actually gone
            var listResult = await _sheetCommands.ListAsync(batch);
            Assert.True(listResult.Success);
            Assert.DoesNotContain(listResult.Worksheets, w => w.Name == "ToDelete");
        }
    }

    [Fact]
    public async Task Copy_WithValidNames_ReturnsSuccessResult()
    {
        // Arrange
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            await _sheetCommands.CreateAsync(batch, "Source");
            await batch.SaveAsync();
        }

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _sheetCommands.CopyAsync(batch, "Source", "Target");
            await batch.SaveAsync();

            // Assert
            Assert.True(result.Success);

            // Verify both source and target sheets exist
            var listResult = await _sheetCommands.ListAsync(batch);
            Assert.True(listResult.Success);
            Assert.Contains(listResult.Worksheets, w => w.Name == "Source");
            Assert.Contains(listResult.Worksheets, w => w.Name == "Target");
        }
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposed) return;

        if (disposing)
        {
            // Cleanup test directory
            try
            {
                if (Directory.Exists(_tempDir))
                {
                    Directory.Delete(_tempDir, recursive: true);
                }
            }
            catch
            {
                // Ignore cleanup failures
            }
        }

        _disposed = true;
    }
}
