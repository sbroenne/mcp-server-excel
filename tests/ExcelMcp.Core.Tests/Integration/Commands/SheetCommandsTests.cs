using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests for Sheet Core operations.
/// These tests require Excel installation and validate Core worksheet data operations.
/// Tests use Core commands directly (not through CLI wrapper).
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
        }
    }

    [Fact]
    public async Task Write_WithValidCsvData_ReturnsSuccessResult()
    {
        // Arrange
        var csvPath = Path.Combine(_tempDir, "test.csv");
        File.WriteAllText(csvPath, "Name,Age\nJohn,30\nJane,25");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _sheetCommands.WriteAsync(batch, "Sheet1", csvPath);
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success);
    }

    [Fact]
    public async Task Read_AfterWrite_ReturnsData()
    {
        // Arrange
        var csvPath = Path.Combine(_tempDir, "test.csv");
        File.WriteAllText(csvPath, "Name,Age\nJohn,30");
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            await _sheetCommands.WriteAsync(batch, "Sheet1", csvPath);
            await batch.SaveAsync();
        }

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _sheetCommands.ReadAsync(batch, "Sheet1", "A1:B2");

            // Assert
            Assert.True(result.Success);
            Assert.NotNull(result.Data);
            Assert.NotEmpty(result.Data);
        }
    }

    [Fact]
    public async Task Clear_WithValidRange_ReturnsSuccessResult()
    {
        // Arrange
        var csvPath = Path.Combine(_tempDir, "test.csv");
        File.WriteAllText(csvPath, "Test,Data\n1,2");
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            await _sheetCommands.WriteAsync(batch, "Sheet1", csvPath);
            await batch.SaveAsync();
        }

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _sheetCommands.ClearAsync(batch, "Sheet1", "A1:B2");
            await batch.SaveAsync();

            // Assert
            Assert.True(result.Success);
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
        }
    }

    [Fact]
    public async Task Append_WithValidData_ReturnsSuccessResult()
    {
        // Arrange
        var csvPath = Path.Combine(_tempDir, "append.csv");
        File.WriteAllText(csvPath, "Name,Value\nTest,123");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _sheetCommands.AppendAsync(batch, "Sheet1", csvPath);
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success);
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
            try
            {
                if (Directory.Exists(_tempDir))
                {
                    Directory.Delete(_tempDir, true);
                }
            }
            catch
            {
                // Ignore cleanup errors
            }
        }

        _disposed = true;
    }
}
