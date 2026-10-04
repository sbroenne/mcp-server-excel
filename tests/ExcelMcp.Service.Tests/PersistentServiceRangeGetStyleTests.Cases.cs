using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceRangeGetStyleTests
{
    [Fact]
    public void GetStyle_UnstyledRange_ReturnsNormalStyle()
    {
        // Arrange & Act
        var batch = _fixture.BatchToken;
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
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set a style first
        RequireSuccess(_commands.SetStyle(batch, sheetName, "A1", "Heading 1"));

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
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set different styles on different cells
        RequireSuccess(_commands.SetStyle(batch, sheetName, "A1", "Heading 1"));
        RequireSuccess(_commands.SetStyle(batch, sheetName, "B1", "Accent1"));
        RequireSuccess(_commands.SetStyle(batch, sheetName, "C1", "Currency"));

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
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set style on entire range (this applies to all cells in the range)
        RequireSuccess(_commands.SetStyle(batch, sheetName, "A1:C3", "Good"));

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
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheetName, "A1", [["retained"]]));
        RequireSuccess(_commands.SetStyle(batch, sheetName, "A1", "Good"));
        var before = PersistentServiceRangeVerification.ReadFormat(_fixture, sheetName, "A1");
        var exception = Assert.Throws<InvalidOperationException>(
            () => _commands.GetStyle(batch, sheetName, "InvalidRange"));

        Assert.Contains("rangeformat.get-style failed [ComInterop/", exception.Message, StringComparison.Ordinal);
        Assert.Equal(before, PersistentServiceRangeVerification.ReadFormat(_fixture, sheetName, "A1"));
        Assert.Equal("Good", RequireSuccess(_commands.GetStyle(batch, sheetName, "A1")).StyleName);
        Assert.Equal("retained", Assert.Single(Assert.Single(
            RequireSuccess(_commands.GetValues(batch, sheetName, "A1")).Values)));
    }
}



