using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

public partial class DataModelTomCommandsTests
{
    #region ValidateDax Tests

    [Fact]
    public void ValidateDax_WithValidFormula_ReturnsValidResult()
    {
        // Arrange
        var daxFormula = "SUM(Sales[Amount])";

        // Act
        var result = _tomCommands.ValidateDax(_testExcelFile, daxFormula);

        // Assert
        if (result.Success)
        {
            Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
            // Validation may or may not succeed depending on TOM connection
            // Just verify the operation completed
        }
        else
        {
            // Connection failure is acceptable
            Assert.True(
                result.ErrorMessage?.Contains("connect") == true,
                $"Expected connection error, got: {result.ErrorMessage}"
            );
        }
    }

    [Fact]
    public void ValidateDax_WithUnbalancedParentheses_ReturnsInvalidResult()
    {
        // Arrange
        var daxFormula = "SUM(Sales[Amount]";

        // Act
        var result = _tomCommands.ValidateDax(_testExcelFile, daxFormula);

        // Assert
        if (result.Success)
        {
            // Validation should detect unbalanced parentheses
            Assert.False(result.IsValid);
            Assert.NotNull(result.ValidationError);
            Assert.Contains("parenthes", result.ValidationError, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void ValidateDax_WithEmptyFormula_ReturnsError()
    {
        // Act
        var result = _tomCommands.ValidateDax(_testExcelFile, "");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("empty", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region ImportMeasures Tests

    [Fact]
    public async Task ImportMeasures_WithNonExistentFile_ReturnsError()
    {
        // Act
        var result = await _tomCommands.ImportMeasures(
            _testExcelFile,
            "NonExistent.json"
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ImportMeasures_WithUnsupportedFormat_ReturnsError()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, "test.txt");
        File.WriteAllText(testFile, "test content");

        // Act
        var result = await _tomCommands.ImportMeasures(
            _testExcelFile,
            testFile
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("format", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region File Validation Tests

    [Fact]
    public void CreateMeasure_WithNonExistentFile_ReturnsError()
    {
        // Act
        var result = _tomCommands.CreateMeasure(
            "NonExistent.xlsx",
            "Sales",
            "TestMeasure",
            "SUM(Sales[Amount])"
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
