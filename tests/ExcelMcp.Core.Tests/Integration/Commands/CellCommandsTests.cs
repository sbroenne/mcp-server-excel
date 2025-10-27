using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests for Cell Core operations using Excel COM automation.
/// Tests Core layer directly (not through CLI wrapper).
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Feature", "Cells")]
[Trait("RequiresExcel", "true")]
public class CoreCellCommandsTests : IDisposable
{
    private readonly ICellCommands _cellCommands;
    private readonly IFileCommands _fileCommands;
    private readonly string _testExcelFile;
    private readonly string _tempDir;

    public CoreCellCommandsTests()
    {
        _cellCommands = new CellCommands();
        _fileCommands = new FileCommands();

        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_CellTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _testExcelFile = Path.Combine(_tempDir, "TestWorkbook.xlsx");

        // Create test Excel file
        var result = _fileCommands.CreateEmptyAsync(_testExcelFile).GetAwaiter().GetResult();
        
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}");
        }
    }

    [Fact]
    public async Task GetValue_WithValidCell_ReturnsSuccess()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _cellCommands.GetValueAsync(batch, "Sheet1", "A1");

        // Assert
        Assert.True(result.Success);
        // Empty cells should return success but may have null/empty value
        Assert.Null(result.ErrorMessage);
    }

    [Fact]
    public async Task SetValue_WithValidCell_ReturnsSuccess()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _cellCommands.SetValueAsync(batch, "Sheet1", "A1", "Test Value");
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success);
        Assert.Null(result.ErrorMessage);
    }

    [Fact]
    public async Task SetValue_ThenGetValue_ReturnsSetValue()
    {
        // Arrange
        string testValue = "Integration Test";

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var setResult = await _cellCommands.SetValueAsync(batch, "Sheet1", "B2", testValue);
            Assert.True(setResult.Success);
            await batch.SaveAsync();
        }
        
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var getResult = await _cellCommands.GetValueAsync(batch, "Sheet1", "B2");
            
            // Assert
            Assert.True(getResult.Success);
            Assert.Equal(testValue, getResult.Value?.ToString());
        }
    }

    [Fact]
    public async Task GetFormula_WithValidCell_ReturnsSuccess()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _cellCommands.GetFormulaAsync(batch, "Sheet1", "C1");

        // Assert
        Assert.True(result.Success);
    }

    [Fact]
    public async Task SetFormula_WithValidFormula_ReturnsSuccess()
    {
        // Arrange
        string formula = "=1+1";

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _cellCommands.SetFormulaAsync(batch, "Sheet1", "D1", formula);
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success);
    }

    [Fact]
    public async Task SetFormula_ThenGetFormula_ReturnsSetFormula()
    {
        // Arrange
        string formula = "=SUM(A1:A10)";

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var setResult = await _cellCommands.SetFormulaAsync(batch, "Sheet1", "E1", formula);
            Assert.True(setResult.Success);
            await batch.SaveAsync();
        }
        
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var getResult = await _cellCommands.GetFormulaAsync(batch, "Sheet1", "E1");
            
            // Assert
            Assert.True(getResult.Success);
            Assert.Equal(formula, getResult.Formula);
        }
    }

    [Fact]
    public async Task GetValue_WithNonExistentFile_ReturnsError()
    {
        // Act & Assert - Should throw when trying to open non-existent file
        await Assert.ThrowsAsync<FileNotFoundException>(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync("nonexistent.xlsx");
        });
    }

    [Fact]
    public async Task SetValue_WithNonExistentFile_ReturnsError()
    {
        // Act & Assert - Should throw when trying to open non-existent file
        await Assert.ThrowsAsync<FileNotFoundException>(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync("nonexistent.xlsx");
        });
    }

    public void Dispose()
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

        GC.SuppressFinalize(this);
    }
}
