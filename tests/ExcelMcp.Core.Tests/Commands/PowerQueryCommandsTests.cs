using Xunit;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using System.IO;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests for Power Query Core operations.
/// These tests require Excel installation and validate Core Power Query data operations.
/// Tests use Core commands directly (not through CLI wrapper).
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PowerQuery")]
public class PowerQueryCommandsTests : IDisposable
{
    private readonly IPowerQueryCommands _powerQueryCommands;
    private readonly IFileCommands _fileCommands;
    private readonly string _testExcelFile;
    private readonly string _testQueryFile;
    private readonly string _tempDir;
    private bool _disposed;

    public PowerQueryCommandsTests()
    {
        _powerQueryCommands = new PowerQueryCommands();
        _fileCommands = new FileCommands();
        
        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_PQ_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        
        _testExcelFile = Path.Combine(_tempDir, "TestWorkbook.xlsx");
        _testQueryFile = Path.Combine(_tempDir, "TestQuery.pq");
        
        // Create test Excel file and Power Query
        CreateTestExcelFile();
        CreateTestQueryFile();
    }

    private void CreateTestExcelFile()
    {
        var result = _fileCommands.CreateEmpty(_testExcelFile, overwriteIfExists: false);
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}. Excel may not be installed.");
        }
    }

    private void CreateTestQueryFile()
    {
        // Create a simple Power Query M file
        string mCode = @"let
    Source = Excel.CurrentWorkbook(){[Name=""Sheet1""]}[Content]
in
    Source";
    
        File.WriteAllText(_testQueryFile, mCode);
    }

    [Fact]
    public void List_WithValidFile_ReturnsSuccessResult()
    {
        // Act
        var result = _powerQueryCommands.List(_testExcelFile);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Queries);
        Assert.Empty(result.Queries); // New file has no queries
    }

    [Fact]
    public async Task Import_WithValidMCode_ReturnsSuccessResult()
    {
        // Act
        var result = await _powerQueryCommands.Import(_testExcelFile, "TestQuery", _testQueryFile);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
    }

    [Fact]
    public async Task List_AfterImport_ShowsNewQuery()
    {
        // Arrange
        await _powerQueryCommands.Import(_testExcelFile, "TestQuery", _testQueryFile);

        // Act
        var result = _powerQueryCommands.List(_testExcelFile);

        // Assert
        Assert.True(result.Success);
        Assert.NotNull(result.Queries);
        Assert.Single(result.Queries);
        Assert.Equal("TestQuery", result.Queries[0].Name);
    }

    [Fact]
    public async Task View_WithExistingQuery_ReturnsMCode()
    {
        // Arrange
        await _powerQueryCommands.Import(_testExcelFile, "TestQuery", _testQueryFile);

        // Act
        var result = _powerQueryCommands.View(_testExcelFile, "TestQuery");

        // Assert
        Assert.True(result.Success);
        Assert.NotNull(result.MCode);
        Assert.Contains("Source", result.MCode);
    }

    [Fact]
    public async Task Export_WithExistingQuery_CreatesFile()
    {
        // Arrange
        await _powerQueryCommands.Import(_testExcelFile, "TestQuery", _testQueryFile);
        var exportPath = Path.Combine(_tempDir, "exported.pq");

        // Act
        var result = await _powerQueryCommands.Export(_testExcelFile, "TestQuery", exportPath);

        // Assert
        Assert.True(result.Success);
        Assert.True(File.Exists(exportPath));
    }

    [Fact]
    public async Task Update_WithValidMCode_ReturnsSuccessResult()
    {
        // Arrange
        await _powerQueryCommands.Import(_testExcelFile, "TestQuery", _testQueryFile);
        var updateFile = Path.Combine(_tempDir, "updated.pq");
        File.WriteAllText(updateFile, "let\n    UpdatedSource = 1\nin\n    UpdatedSource");

        // Act
        var result = await _powerQueryCommands.Update(_testExcelFile, "TestQuery", updateFile);

        // Assert
        Assert.True(result.Success);
    }

    [Fact]
    public async Task Delete_WithExistingQuery_ReturnsSuccessResult()
    {
        // Arrange
        await _powerQueryCommands.Import(_testExcelFile, "TestQuery", _testQueryFile);

        // Act
        var result = _powerQueryCommands.Delete(_testExcelFile, "TestQuery");

        // Assert
        Assert.True(result.Success);
    }

    [Fact]
    public void View_WithNonExistentQuery_ReturnsErrorResult()
    {
        // Act
        var result = _powerQueryCommands.View(_testExcelFile, "NonExistentQuery");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public async Task Import_ThenDelete_ThenList_ShowsEmpty()
    {
        // Arrange
        await _powerQueryCommands.Import(_testExcelFile, "TestQuery", _testQueryFile);
        _powerQueryCommands.Delete(_testExcelFile, "TestQuery");

        // Act
        var result = _powerQueryCommands.List(_testExcelFile);

        // Assert
        Assert.True(result.Success);
        Assert.Empty(result.Queries);
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
