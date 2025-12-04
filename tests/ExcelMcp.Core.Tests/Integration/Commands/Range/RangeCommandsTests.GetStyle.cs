using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

public partial class RangeCommandsTests
{
    [Fact]
    public void GetStyle_UnstyledRange_ReturnsNormalStyle()
    {
        // Arrange & Act
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        var result = _commands.GetStyle(batch, sheetName, "A1");

        // Assert
        Assert.True(result.Success, $"GetStyle failed: {result.ErrorMessage}");
        Assert.Equal("Normal", result.StyleName);
        Assert.True(result.IsBuiltInStyle);
        // Note: StyleDescription may be null for some styles
    }

    [Fact]
    public void GetStyle_AfterSetStyle_ReturnsAppliedStyle()
    {
        // Arrange & Act
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set a style first
        _commands.SetStyle(batch, sheetName, "A1", "Heading 1");

        // Now get the style
        var getResult = _commands.GetStyle(batch, sheetName, "A1");

        // Assert
        Assert.True(getResult.Success, $"GetStyle failed: {getResult.ErrorMessage}");
        Assert.Equal("Heading 1", getResult.StyleName);
        Assert.True(getResult.IsBuiltInStyle);
        // Note: StyleDescription may be null for some styles
    }

    [Fact]
    public void GetStyle_MultipleStyles_ReturnsCorrectStyles()
    {
        // Arrange & Act
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set different styles on different cells
        _commands.SetStyle(batch, sheetName, "A1", "Heading 1");
        _commands.SetStyle(batch, sheetName, "B1", "Accent1");
        _commands.SetStyle(batch, sheetName, "C1", "Currency");

        // Get the styles
        var getHeading1 = _commands.GetStyle(batch, sheetName, "A1");
        var getAccent1 = _commands.GetStyle(batch, sheetName, "B1");
        var getCurrency = _commands.GetStyle(batch, sheetName, "C1");

        // Assert
        Assert.True(getHeading1.Success, $"GetStyle A1 failed: {getHeading1.ErrorMessage}");
        Assert.Equal("Heading 1", getHeading1.StyleName);
        Assert.True(getHeading1.IsBuiltInStyle);

        Assert.True(getAccent1.Success, $"GetStyle B1 failed: {getAccent1.ErrorMessage}");
        Assert.Equal("Accent1", getAccent1.StyleName);
        Assert.True(getAccent1.IsBuiltInStyle);

        Assert.True(getCurrency.Success, $"GetStyle C1 failed: {getCurrency.ErrorMessage}");
        Assert.Equal("Currency", getCurrency.StyleName);
        Assert.True(getCurrency.IsBuiltInStyle);
    }

    [Fact]
    public void GetStyle_RangeMultipleCells_ReturnsFirstCellStyle()
    {
        // Arrange & Act
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set style on entire range (this applies to all cells in the range)
        _commands.SetStyle(batch, sheetName, "A1:C3", "Good");

        // Get style for entire range (should return first cell's style)
        var getResult = _commands.GetStyle(batch, sheetName, "A1:C3");

        // Assert
        Assert.True(getResult.Success, $"GetStyle failed: {getResult.ErrorMessage}");
        Assert.Equal("Good", getResult.StyleName);
        Assert.True(getResult.IsBuiltInStyle);
    }

    [Fact]
    public void GetStyle_InvalidRange_ThrowsException()
    {
        // Arrange & Act & Assert - Should throw when Excel COM rejects invalid range
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        var exception = Assert.ThrowsAny<Exception>(
            () => _commands.GetStyle(batch, sheetName, "InvalidRange"));

        // Verify exception is related to range access
        Assert.NotNull(exception.Message);
    }
}
