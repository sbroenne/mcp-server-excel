using Sbroenne.ExcelMcp.Core.Commands;
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
        var result = _fileCommands.CreateEmpty(_testExcelFile, overwriteIfExists: false);
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}");
        }
    }

    [Fact]
    public void List_WithValidFile_ReturnsSuccessResult()
    {
        // Act
        var result = _sheetCommands.List(_testExcelFile);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Worksheets);
        Assert.NotEmpty(result.Worksheets); // New Excel file has Sheet1
    }

    [Fact]
    public void Create_WithValidName_ReturnsSuccessResult()
    {
        // Act
        var result = _sheetCommands.Create(_testExcelFile, "TestSheet");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
    }

    [Fact]
    public void List_AfterCreate_ShowsNewSheet()
    {
        // Arrange
        _sheetCommands.Create(_testExcelFile, "TestSheet");

        // Act
        var result = _sheetCommands.List(_testExcelFile);

        // Assert
        Assert.True(result.Success);
        Assert.Contains(result.Worksheets, w => w.Name == "TestSheet");
    }

    [Fact]
    public void Rename_WithValidNames_ReturnsSuccessResult()
    {
        // Arrange
        _sheetCommands.Create(_testExcelFile, "OldName");

        // Act
        var result = _sheetCommands.Rename(_testExcelFile, "OldName", "NewName");

        // Assert
        Assert.True(result.Success);
    }

    [Fact]
    public void Delete_WithExistingSheet_ReturnsSuccessResult()
    {
        // Arrange
        _sheetCommands.Create(_testExcelFile, "ToDelete");

        // Act
        var result = _sheetCommands.Delete(_testExcelFile, "ToDelete");

        // Assert
        Assert.True(result.Success);
    }

    [Fact]
    public void Write_WithValidCsvData_ReturnsSuccessResult()
    {
        // Arrange
        var csvPath = Path.Combine(_tempDir, "test.csv");
        File.WriteAllText(csvPath, "Name,Age\nJohn,30\nJane,25");

        // Act
        var result = _sheetCommands.Write(_testExcelFile, "Sheet1", csvPath);

        // Assert
        Assert.True(result.Success);
    }

    [Fact]
    public void Read_AfterWrite_ReturnsData()
    {
        // Arrange
        var csvPath = Path.Combine(_tempDir, "test.csv");
        File.WriteAllText(csvPath, "Name,Age\nJohn,30");
        _sheetCommands.Write(_testExcelFile, "Sheet1", csvPath);

        // Act
        var result = _sheetCommands.Read(_testExcelFile, "Sheet1", "A1:B2");

        // Assert
        Assert.True(result.Success);
        Assert.NotNull(result.Data);
        Assert.NotEmpty(result.Data);
    }

    [Fact]
    public void Clear_WithValidRange_ReturnsSuccessResult()
    {
        // Arrange
        var csvPath = Path.Combine(_tempDir, "test.csv");
        File.WriteAllText(csvPath, "Test,Data\n1,2");
        _sheetCommands.Write(_testExcelFile, "Sheet1", csvPath);

        // Act
        var result = _sheetCommands.Clear(_testExcelFile, "Sheet1", "A1:B2");

        // Assert
        Assert.True(result.Success);
    }

    [Fact]
    public void Copy_WithValidNames_ReturnsSuccessResult()
    {
        // Arrange
        _sheetCommands.Create(_testExcelFile, "Source");

        // Act
        var result = _sheetCommands.Copy(_testExcelFile, "Source", "Target");

        // Assert
        Assert.True(result.Success);
    }

    [Fact]
    public void Append_WithValidData_ReturnsSuccessResult()
    {
        // Arrange
        var csvPath = Path.Combine(_tempDir, "append.csv");
        File.WriteAllText(csvPath, "Name,Value\nTest,123");

        // Act
        var result = _sheetCommands.Append(_testExcelFile, "Sheet1", csvPath);

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
