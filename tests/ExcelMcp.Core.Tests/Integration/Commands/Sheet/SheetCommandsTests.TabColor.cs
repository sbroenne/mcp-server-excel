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
    public async Task SetTabColor_WithValidRGB_SetsColorCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests),
            nameof(SetTabColor_WithValidRGB_SetsColorCorrectly),
            _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        await _sheetCommands.CreateAsync(batch, "ColorTest");

        // Act - Set red color
        var setResult = await _sheetCommands.SetTabColorAsync(batch, "ColorTest", 255, 0, 0);

        // Assert - Verify set succeeded
        Assert.True(setResult.Success, $"SetTabColor failed: {setResult.ErrorMessage}");

        // Verify color was actually set by reading it back
        var getResult = await _sheetCommands.GetTabColorAsync(batch, "ColorTest");
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
    public async Task SetTabColor_WithDifferentColors_AllSetCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests),
            nameof(SetTabColor_WithDifferentColors_AllSetCorrectly),
            _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Create multiple sheets
        await _sheetCommands.CreateAsync(batch, "Red");
        await _sheetCommands.CreateAsync(batch, "Green");
        await _sheetCommands.CreateAsync(batch, "Blue");

        // Act - Set different colors
        await _sheetCommands.SetTabColorAsync(batch, "Red", 255, 0, 0);
        await _sheetCommands.SetTabColorAsync(batch, "Green", 0, 255, 0);
        await _sheetCommands.SetTabColorAsync(batch, "Blue", 0, 0, 255);

        // Assert - Verify each color
        var redColor = await _sheetCommands.GetTabColorAsync(batch, "Red");
        Assert.True(redColor.HasColor);
        Assert.Equal(255, redColor.Red);
        Assert.Equal(0, redColor.Green);
        Assert.Equal(0, redColor.Blue);

        var greenColor = await _sheetCommands.GetTabColorAsync(batch, "Green");
        Assert.True(greenColor.HasColor);
        Assert.Equal(0, greenColor.Red);
        Assert.Equal(255, greenColor.Green);
        Assert.Equal(0, greenColor.Blue);

        var blueColor = await _sheetCommands.GetTabColorAsync(batch, "Blue");
        Assert.True(blueColor.HasColor);
        Assert.Equal(0, blueColor.Red);
        Assert.Equal(0, blueColor.Green);
        Assert.Equal(255, blueColor.Blue);

        // Save changes
    }
    /// <inheritdoc/>

    [Fact]
    public async Task GetTabColor_WithNoColor_ReturnsHasColorFalse()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests),
            nameof(GetTabColor_WithNoColor_ReturnsHasColorFalse),
            _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        await _sheetCommands.CreateAsync(batch, "NoColor");

        // Act
        var result = await _sheetCommands.GetTabColorAsync(batch, "NoColor");

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
    public async Task ClearTabColor_RemovesColor()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests),
            nameof(ClearTabColor_RemovesColor),
            _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        await _sheetCommands.CreateAsync(batch, "ClearTest");
        await _sheetCommands.SetTabColorAsync(batch, "ClearTest", 255, 165, 0); // Orange

        // Verify color is set
        var beforeClear = await _sheetCommands.GetTabColorAsync(batch, "ClearTest");
        Assert.True(beforeClear.HasColor);

        // Act - Clear color
        var clearResult = await _sheetCommands.ClearTabColorAsync(batch, "ClearTest");

        // Assert
        Assert.True(clearResult.Success);

        var afterClear = await _sheetCommands.GetTabColorAsync(batch, "ClearTest");
        Assert.True(afterClear.Success);
        Assert.False(afterClear.HasColor);

        // Save changes
    }
    /// <inheritdoc/>

    [Fact]
    public async Task SetTabColor_WithInvalidRGB_ReturnsError()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests),
            nameof(SetTabColor_WithInvalidRGB_ReturnsError),
            _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        await _sheetCommands.CreateAsync(batch, "InvalidColor");

        // Act - Try to set invalid RGB values
        var result1 = await _sheetCommands.SetTabColorAsync(batch, "InvalidColor", 256, 0, 0); // Red too high
        var result2 = await _sheetCommands.SetTabColorAsync(batch, "InvalidColor", 0, -1, 0); // Green negative

        // Assert
        Assert.False(result1.Success);
        Assert.Contains("must be between 0 and 255", result1.ErrorMessage);

        Assert.False(result2.Success);
        Assert.Contains("must be between 0 and 255", result2.ErrorMessage);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task SetTabColor_WithNonExistentSheet_ReturnsError()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests),
            nameof(SetTabColor_WithNonExistentSheet_ReturnsError),
            _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act
        var result = await _sheetCommands.SetTabColorAsync(batch, "NonExistent", 255, 0, 0);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task TabColor_RGBToBGRConversion_WorksCorrectly()
    {
        // Arrange - Test BGR conversion accuracy
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests),
            nameof(TabColor_RGBToBGRConversion_WorksCorrectly),
            _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        await _sheetCommands.CreateAsync(batch, "ConversionTest");

        // Act - Set a complex color (purple: RGB(128, 0, 128))
        await _sheetCommands.SetTabColorAsync(batch, "ConversionTest", 128, 0, 128);

        // Assert - Verify conversion accuracy
        var result = await _sheetCommands.GetTabColorAsync(batch, "ConversionTest");
        Assert.True(result.Success);
        Assert.True(result.HasColor);
        Assert.Equal(128, result.Red);
        Assert.Equal(0, result.Green);
        Assert.Equal(128, result.Blue);
        Assert.Equal("#800080", result.HexColor);

        // Save changes
    }
}
