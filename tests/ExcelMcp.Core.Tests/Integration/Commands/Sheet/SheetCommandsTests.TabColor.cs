using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Sheet;

/// <summary>
/// Integration tests for worksheet tab color operations
/// </summary>
public partial class SheetCommandsTests
{
    /// <inheritdoc/>

    [Fact]
    public void SetTabColor_WithValidRGB_SetsColorCorrectly()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests),
            nameof(SetTabColor_WithValidRGB_SetsColorCorrectly),
            _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "ColorTest");

        // Act - Set red color
        var setResult = _sheetCommands.SetTabColor(batch, "ColorTest", 255, 0, 0);

        // Assert - Verify set succeeded
        Assert.True(setResult.Success, $"SetTabColor failed: {setResult.ErrorMessage}");

        // Verify color was actually set by reading it back
        var getResult = _sheetCommands.GetTabColor(batch, "ColorTest");
        Assert.True(getResult.Success);
        Assert.True(getResult.HasColor);
        Assert.Equal(255, getResult.Red);
        Assert.Equal(0, getResult.Green);
        Assert.Equal(0, getResult.Blue);
        Assert.Equal("#FF0000", getResult.HexColor);

        // Save changes
    }
    /// <inheritdoc/>

    [Fact]
    public void SetTabColor_WithDifferentColors_AllSetCorrectly()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests),
            nameof(SetTabColor_WithDifferentColors_AllSetCorrectly),
            _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create multiple sheets
        _sheetCommands.Create(batch, "Red");
        _sheetCommands.Create(batch, "Green");
        _sheetCommands.Create(batch, "Blue");

        // Act - Set different colors
        _sheetCommands.SetTabColor(batch, "Red", 255, 0, 0);
        _sheetCommands.SetTabColor(batch, "Green", 0, 255, 0);
        _sheetCommands.SetTabColor(batch, "Blue", 0, 0, 255);

        // Assert - Verify each color
        var redColor = _sheetCommands.GetTabColor(batch, "Red");
        Assert.True(redColor.HasColor);
        Assert.Equal(255, redColor.Red);
        Assert.Equal(0, redColor.Green);
        Assert.Equal(0, redColor.Blue);

        var greenColor = _sheetCommands.GetTabColor(batch, "Green");
        Assert.True(greenColor.HasColor);
        Assert.Equal(0, greenColor.Red);
        Assert.Equal(255, greenColor.Green);
        Assert.Equal(0, greenColor.Blue);

        var blueColor = _sheetCommands.GetTabColor(batch, "Blue");
        Assert.True(blueColor.HasColor);
        Assert.Equal(0, blueColor.Red);
        Assert.Equal(0, blueColor.Green);
        Assert.Equal(255, blueColor.Blue);

        // Save changes
    }
    /// <inheritdoc/>

    [Fact]
    public void GetTabColor_WithNoColor_ReturnsHasColorFalse()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests),
            nameof(GetTabColor_WithNoColor_ReturnsHasColorFalse),
            _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "NoColor");

        // Act
        var result = _sheetCommands.GetTabColor(batch, "NoColor");

        // Assert
        Assert.True(result.Success);
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
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests),
            nameof(ClearTabColor_RemovesColor),
            _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "ClearTest");
        _sheetCommands.SetTabColor(batch, "ClearTest", 255, 165, 0); // Orange

        // Verify color is set
        var beforeClear = _sheetCommands.GetTabColor(batch, "ClearTest");
        Assert.True(beforeClear.HasColor);

        // Act - Clear color
        var clearResult = _sheetCommands.ClearTabColor(batch, "ClearTest");

        // Assert
        Assert.True(clearResult.Success);

        var afterClear = _sheetCommands.GetTabColor(batch, "ClearTest");
        Assert.True(afterClear.Success);
        Assert.False(afterClear.HasColor);

        // Save changes
    }
    /// <inheritdoc/>

    [Fact]
    public void SetTabColor_WithInvalidRGB_ReturnsError()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests),
            nameof(SetTabColor_WithInvalidRGB_ReturnsError),
            _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "InvalidColor");

        // Act - Try to set invalid RGB values
        var result1 = _sheetCommands.SetTabColor(batch, "InvalidColor", 256, 0, 0); // Red too high
        var result2 = _sheetCommands.SetTabColor(batch, "InvalidColor", 0, -1, 0); // Green negative

        // Assert
        Assert.False(result1.Success);
        Assert.Contains("must be between 0 and 255", result1.ErrorMessage);

        Assert.False(result2.Success);
        Assert.Contains("must be between 0 and 255", result2.ErrorMessage);
    }
    /// <inheritdoc/>

    [Fact]
    public void SetTabColor_WithNonExistentSheet_ReturnsError()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests),
            nameof(SetTabColor_WithNonExistentSheet_ReturnsError),
            _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act & Assert - Non-existent sheet should throw InvalidOperationException
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _sheetCommands.SetTabColor(batch, "NonExistent", 255, 0, 0));

        Assert.Contains("not found", exception.Message);
    }
    /// <inheritdoc/>

    [Fact]
    public void TabColor_RGBToBGRConversion_WorksCorrectly()
    {
        // Arrange - Test BGR conversion accuracy
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests),
            nameof(TabColor_RGBToBGRConversion_WorksCorrectly),
            _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "ConversionTest");

        // Act - Set a complex color (purple: RGB(128, 0, 128))
        _sheetCommands.SetTabColor(batch, "ConversionTest", 128, 0, 128);

        // Assert - Verify conversion accuracy
        var result = _sheetCommands.GetTabColor(batch, "ConversionTest");
        Assert.True(result.Success);
        Assert.True(result.HasColor);
        Assert.Equal(128, result.Red);
        Assert.Equal(0, result.Green);
        Assert.Equal(128, result.Blue);
        Assert.Equal("#800080", result.HexColor);

        // Save changes
    }
}
