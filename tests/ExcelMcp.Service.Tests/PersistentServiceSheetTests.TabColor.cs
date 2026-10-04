using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for worksheet tab color operations
/// </summary>
public sealed partial class PersistentServiceSheetTests
{
    /// <inheritdoc/>

    [Fact]
    public void SetTabColor_WithValidRGB_SetsColorCorrectly()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = $"Color_{Guid.NewGuid():N}"[..31];
        _fixture.CreateNamedTestSheet(batch, sheetName);

        // Act - Set red color
        _sheetCommands.SetTabColor(batch, sheetName, 255, 0, 0);  // SetTabColor throws on error

        // Assert - reaching here means set succeeded

        // Verify color was actually set by reading it back
        var getResult = _sheetCommands.GetTabColor(batch, sheetName);
        RequireSuccess(getResult);
        Assert.True(getResult.HasColor);
        Assert.Equal(255, getResult.Red);
        Assert.Equal(0, getResult.Green);
        Assert.Equal(0, getResult.Blue);
        Assert.Equal("#FF0000", getResult.HexColor);
    }
    /// <inheritdoc/>

    [Fact]
    public void SetTabColor_WithDifferentColors_AllSetCorrectly()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var uniqueId = Guid.NewGuid().ToString("N")[..6];
        var redSheet = $"Red_{uniqueId}";
        var greenSheet = $"Green_{uniqueId}";
        var blueSheet = $"Blue_{uniqueId}";

        // Create multiple sheets
        _fixture.CreateNamedTestSheet(batch, redSheet);
        _fixture.CreateNamedTestSheet(batch, greenSheet);
        _fixture.CreateNamedTestSheet(batch, blueSheet);

        // Act - Set different colors
        _sheetCommands.SetTabColor(batch, redSheet, 255, 0, 0);
        _sheetCommands.SetTabColor(batch, greenSheet, 0, 255, 0);
        _sheetCommands.SetTabColor(batch, blueSheet, 0, 0, 255);

        // Assert - Verify each color
        var redColor = _sheetCommands.GetTabColor(batch, redSheet);
        RequireSuccess(redColor);
        Assert.True(redColor.HasColor);
        Assert.Equal(255, redColor.Red);
        Assert.Equal(0, redColor.Green);
        Assert.Equal(0, redColor.Blue);

        var greenColor = _sheetCommands.GetTabColor(batch, greenSheet);
        RequireSuccess(greenColor);
        Assert.True(greenColor.HasColor);
        Assert.Equal(0, greenColor.Red);
        Assert.Equal(255, greenColor.Green);
        Assert.Equal(0, greenColor.Blue);

        var blueColor = _sheetCommands.GetTabColor(batch, blueSheet);
        RequireSuccess(blueColor);
        Assert.True(blueColor.HasColor);
        Assert.Equal(0, blueColor.Red);
        Assert.Equal(0, blueColor.Green);
        Assert.Equal(255, blueColor.Blue);
    }
    /// <inheritdoc/>

    [Fact]
    public void GetTabColor_WithNoColor_ReturnsHasColorFalse()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = $"NoClr_{Guid.NewGuid():N}"[..31];
        _fixture.CreateNamedTestSheet(batch, sheetName);

        // Act
        var result = _sheetCommands.GetTabColor(batch, sheetName);

        // Assert
        RequireSuccess(result);
        Assert.False(result.HasColor);
        Assert.Null(result.Red);
        Assert.Null(result.Green);
        Assert.Null(result.Blue);
        Assert.Null(result.HexColor);
    }
    /// <inheritdoc/>

    [Fact]
    public void ClearTabColor_RemovesColor()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = $"Clear_{Guid.NewGuid():N}"[..31];
        _fixture.CreateNamedTestSheet(batch, sheetName);
        _sheetCommands.SetTabColor(batch, sheetName, 255, 165, 0); // Orange

        // Verify color is set
        var beforeClear = _sheetCommands.GetTabColor(batch, sheetName);
        RequireSuccess(beforeClear);
        Assert.True(beforeClear.HasColor);
        Assert.Equal("#FFA500", beforeClear.HexColor);
        var untouchedName = _fixture.CreateTestSheet(batch);
        _sheetCommands.SetTabColor(batch, untouchedName, 12, 34, 56);

        // Act - Clear color
        _sheetCommands.ClearTabColor(batch, sheetName);  // ClearTabColor throws on error

        // Assert - reaching here means clear succeeded

        var afterClear = _sheetCommands.GetTabColor(batch, sheetName);
        RequireSuccess(afterClear);
        Assert.False(afterClear.HasColor);
        Assert.Null(afterClear.HexColor);
        Assert.Null(afterClear.Red);
        Assert.Null(afterClear.Green);
        Assert.Null(afterClear.Blue);
        var untouched = _sheetCommands.GetTabColor(batch, untouchedName);
        RequireSuccess(untouched);
        Assert.Equal("#0C2238", untouched.HexColor);
    }
    /// <inheritdoc/>

    [Fact]
    public void SetTabColor_WithInvalidRGB_ThrowsException()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = $"Inv_{Guid.NewGuid():N}"[..31];
        _fixture.CreateNamedTestSheet(batch, sheetName);
        _sheetCommands.SetTabColor(batch, sheetName, 12, 34, 56);
        var before = _sheetCommands.GetTabColor(batch, sheetName);
        RequireSuccess(before);
        Assert.Equal("#0C2238", before.HexColor);

        // Act & Assert - Should throw ArgumentException for invalid RGB values
        var exception1 = Assert.Throws<ArgumentException>(
            () => _sheetCommands.SetTabColor(batch, sheetName, 256, 0, 0)); // Red too high
        Assert.Contains("must be between 0 and 255", exception1.Message);
        AssertTabColorPreserved();

        var exception2 = Assert.Throws<ArgumentException>(
            () => _sheetCommands.SetTabColor(batch, sheetName, 0, -1, 0)); // Green negative
        Assert.Contains("must be between 0 and 255", exception2.Message);
        AssertTabColorPreserved();
        var exception3 = Assert.Throws<ArgumentException>(
            () => _sheetCommands.SetTabColor(batch, sheetName, 0, 0, 256));
        Assert.Contains("must be between 0 and 255", exception3.Message);
        AssertTabColorPreserved();

        void AssertTabColorPreserved()
        {
            var after = _sheetCommands.GetTabColor(batch, sheetName);
            RequireSuccess(after);
            Assert.True(after.HasColor);
            Assert.Equal(before.Red, after.Red);
            Assert.Equal(before.Green, after.Green);
            Assert.Equal(before.Blue, after.Blue);
            Assert.Equal(before.HexColor, after.HexColor);
        }
    }
    /// <inheritdoc/>

    [Fact]
    public void SetTabColor_WithNonExistentSheet_ThrowsException()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Act & Assert - Non-existent sheet should throw InvalidOperationException
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _sheetCommands.SetTabColor(batch, $"NonExistent_{Guid.NewGuid():N}", 255, 0, 0));

        Assert.Contains("not found", exception.Message);
    }
    /// <inheritdoc/>

    [Fact]
    public void TabColor_RGBToBGRConversion_WorksCorrectly()
    {
        // Arrange - Test BGR conversion accuracy
        var batch = _fixture.BatchToken;
        var sheetName = $"Conv_{Guid.NewGuid():N}"[..31];
        _fixture.CreateNamedTestSheet(batch, sheetName);

        // Act - Set a complex color (purple: RGB(128, 0, 128))
        _sheetCommands.SetTabColor(batch, sheetName, 128, 0, 128);

        // Assert - Verify conversion accuracy
        var result = _sheetCommands.GetTabColor(batch, sheetName);
        RequireSuccess(result);
        Assert.True(result.HasColor);
        Assert.Equal(128, result.Red);
        Assert.Equal(0, result.Green);
        Assert.Equal(128, result.Blue);
        Assert.Equal("#800080", result.HexColor);
    }
}



