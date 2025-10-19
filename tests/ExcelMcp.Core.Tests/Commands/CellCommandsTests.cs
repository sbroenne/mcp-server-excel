using Xunit;
using Sbroenne.ExcelMcp.Core.Commands;
using System.IO;

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
public class CellCommandsTests : IDisposable
{
    private readonly ICellCommands _cellCommands;
    private readonly IFileCommands _fileCommands;
    private readonly string _testExcelFile;
    private readonly string _tempDir;

    public CellCommandsTests()
    {
        _cellCommands = new CellCommands();
        _fileCommands = new FileCommands();
        
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_CellTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        
        _testExcelFile = Path.Combine(_tempDir, "TestWorkbook.xlsx");
        
        // Create test Excel file
        var result = _fileCommands.CreateEmpty(_testExcelFile);
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}");
        }
    }

    [Fact]
    public void GetValue_WithValidCell_ReturnsSuccess()
    {
        // Act
        var result = _cellCommands.GetValue(_testExcelFile, "Sheet1", "A1");

        // Assert
        Assert.True(result.Success);
        Assert.NotNull(result.Value);
    }

    [Fact]
    public void SetValue_WithValidCell_ReturnsSuccess()
    {
        // Act
        var result = _cellCommands.SetValue(_testExcelFile, "Sheet1", "A1", "Test Value");

        // Assert
        Assert.True(result.Success);
        Assert.Null(result.ErrorMessage);
    }

    [Fact]
    public void SetValue_ThenGetValue_ReturnsSetValue()
    {
        // Arrange
        string testValue = "Integration Test";

        // Act
        var setResult = _cellCommands.SetValue(_testExcelFile, "Sheet1", "B2", testValue);
        var getResult = _cellCommands.GetValue(_testExcelFile, "Sheet1", "B2");

        // Assert
        Assert.True(setResult.Success);
        Assert.True(getResult.Success);
        Assert.Equal(testValue, getResult.Value?.ToString());
    }

    [Fact]
    public void GetFormula_WithValidCell_ReturnsSuccess()
    {
        // Act
        var result = _cellCommands.GetFormula(_testExcelFile, "Sheet1", "C1");

        // Assert
        Assert.True(result.Success);
    }

    [Fact]
    public void SetFormula_WithValidFormula_ReturnsSuccess()
    {
        // Arrange
        string formula = "=1+1";

        // Act
        var result = _cellCommands.SetFormula(_testExcelFile, "Sheet1", "D1", formula);

        // Assert
        Assert.True(result.Success);
    }

    [Fact]
    public void SetFormula_ThenGetFormula_ReturnsSetFormula()
    {
        // Arrange
        string formula = "=SUM(A1:A10)";

        // Act
        var setResult = _cellCommands.SetFormula(_testExcelFile, "Sheet1", "E1", formula);
        var getResult = _cellCommands.GetFormula(_testExcelFile, "Sheet1", "E1");

        // Assert
        Assert.True(setResult.Success);
        Assert.True(getResult.Success);
        Assert.Equal(formula, getResult.Formula);
    }

    [Fact]
    public void GetValue_WithNonExistentFile_ReturnsError()
    {
        // Act
        var result = _cellCommands.GetValue("nonexistent.xlsx", "Sheet1", "A1");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void SetValue_WithNonExistentFile_ReturnsError()
    {
        // Act
        var result = _cellCommands.SetValue("nonexistent.xlsx", "Sheet1", "A1", "Value");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
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
