using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Sheet;

/// <summary>
/// Integration tests for worksheet visibility operations
/// </summary>
public partial class SheetCommandsTests
{
    /// <inheritdoc/>

    [Fact]
    public async Task SetVisibility_ToHidden_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests),
            nameof(SetVisibility_ToHidden_WorksCorrectly),
            _tempDir);

        // Use filePath-based API
        await _sheetCommands.CreateAsync(testFile, "HideTest");

        // Act
        var setResult = await _sheetCommands.SetVisibilityAsync(testFile, "HideTest", SheetVisibility.Hidden);

        // Assert
        Assert.True(setResult.Success, $"SetVisibility failed: {setResult.ErrorMessage}");

        // Verify by reading visibility
        var getResult = await _sheetCommands.GetVisibilityAsync(testFile, "HideTest");
        Assert.True(getResult.Success);
        Assert.Equal(SheetVisibility.Hidden, getResult.Visibility);
        Assert.Equal("Hidden", getResult.VisibilityName);

        // Save changes
    }
    /// <inheritdoc/>

    [Fact]
    public async Task SetVisibility_ToVeryHidden_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests),
            nameof(SetVisibility_ToVeryHidden_WorksCorrectly),
            _tempDir);

        // Use filePath-based API
        await _sheetCommands.CreateAsync(testFile, "VeryHideTest");

        // Act
        var setResult = await _sheetCommands.SetVisibilityAsync(testFile, "VeryHideTest", SheetVisibility.VeryHidden);

        // Assert
        Assert.True(setResult.Success);

        var getResult = await _sheetCommands.GetVisibilityAsync(testFile, "VeryHideTest");
        Assert.True(getResult.Success);
        Assert.Equal(SheetVisibility.VeryHidden, getResult.Visibility);
        Assert.Equal("VeryHidden", getResult.VisibilityName);

        // Save changes
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Show_HiddenSheet_MakesVisible()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests),
            nameof(Show_HiddenSheet_MakesVisible),
            _tempDir);

        // Use filePath-based API
        await _sheetCommands.CreateAsync(testFile, "ShowTest");
        await _sheetCommands.HideAsync(testFile, "ShowTest");

        // Verify it's hidden
        var hiddenCheck = await _sheetCommands.GetVisibilityAsync(testFile, "ShowTest");
        Assert.Equal(SheetVisibility.Hidden, hiddenCheck.Visibility);

        // Act - Show the sheet
        var showResult = await _sheetCommands.ShowAsync(testFile, "ShowTest");

        // Assert
        Assert.True(showResult.Success);

        var visibleCheck = await _sheetCommands.GetVisibilityAsync(testFile, "ShowTest");
        Assert.Equal(SheetVisibility.Visible, visibleCheck.Visibility);

        // Save changes
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Show_VeryHiddenSheet_MakesVisible()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests),
            nameof(Show_VeryHiddenSheet_MakesVisible),
            _tempDir);

        // Use filePath-based API
        await _sheetCommands.CreateAsync(testFile, "VeryHideShowTest");
        await _sheetCommands.VeryHideAsync(testFile, "VeryHideShowTest");

        // Verify it's very hidden
        var veryHiddenCheck = await _sheetCommands.GetVisibilityAsync(testFile, "VeryHideShowTest");
        Assert.Equal(SheetVisibility.VeryHidden, veryHiddenCheck.Visibility);

        // Act - Show the sheet
        var showResult = await _sheetCommands.ShowAsync(testFile, "VeryHideShowTest");

        // Assert
        Assert.True(showResult.Success);

        var visibleCheck = await _sheetCommands.GetVisibilityAsync(testFile, "VeryHideShowTest");
        Assert.Equal(SheetVisibility.Visible, visibleCheck.Visibility);

        // Save changes
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Hide_VisibleSheet_MakesHidden()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests),
            nameof(Hide_VisibleSheet_MakesHidden),
            _tempDir);

        // Use filePath-based API
        await _sheetCommands.CreateAsync(testFile, "HideMe");

        // Act
        var hideResult = await _sheetCommands.HideAsync(testFile, "HideMe");

        // Assert
        Assert.True(hideResult.Success);

        var getResult = await _sheetCommands.GetVisibilityAsync(testFile, "HideMe");
        Assert.Equal(SheetVisibility.Hidden, getResult.Visibility);

        // Save changes
    }
    /// <inheritdoc/>

    [Fact]
    public async Task VeryHide_VeryHidesVisibleSheet()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests),
            nameof(VeryHide_VeryHidesVisibleSheet),
            _tempDir);

        // Use filePath-based API
        await _sheetCommands.CreateAsync(testFile, "VeryHideMe");

        // Act
        var veryHideResult = await _sheetCommands.VeryHideAsync(testFile, "VeryHideMe");

        // Assert
        Assert.True(veryHideResult.Success);

        var getResult = await _sheetCommands.GetVisibilityAsync(testFile, "VeryHideMe");
        Assert.Equal(SheetVisibility.VeryHidden, getResult.Visibility);

        // Save changes
    }
    /// <inheritdoc/>

    [Fact]
    public async Task GetVisibility_ForVisibleSheet_ReturnsVisible()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests),
            nameof(GetVisibility_ForVisibleSheet_ReturnsVisible),
            _tempDir);

        // Use filePath-based API
        await _sheetCommands.CreateAsync(testFile, "VisibleSheet");

        // Act
        var result = await _sheetCommands.GetVisibilityAsync(testFile, "VisibleSheet");

        // Assert
        Assert.True(result.Success);
        Assert.Equal(SheetVisibility.Visible, result.Visibility);
        Assert.Equal("Visible", result.VisibilityName);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task SetVisibility_WithNonExistentSheet_ReturnsError()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests),
            nameof(SetVisibility_WithNonExistentSheet_ReturnsError),
            _tempDir);

        // Use filePath-based API

        // Act
        var result = await _sheetCommands.SetVisibilityAsync(testFile, "NonExistent", SheetVisibility.Hidden);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Visibility_CompleteWorkflow_AllLevelsWork()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests),
            nameof(Visibility_CompleteWorkflow_AllLevelsWork),
            _tempDir);

        // Use filePath-based API
        await _sheetCommands.CreateAsync(testFile, "Workflow");

        // Act & Assert - Test complete visibility workflow

        // Start visible
        var check1 = await _sheetCommands.GetVisibilityAsync(testFile, "Workflow");
        Assert.Equal(SheetVisibility.Visible, check1.Visibility);

        // Hide it
        await _sheetCommands.HideAsync(testFile, "Workflow");
        var check2 = await _sheetCommands.GetVisibilityAsync(testFile, "Workflow");
        Assert.Equal(SheetVisibility.Hidden, check2.Visibility);

        // Very hide it
        await _sheetCommands.VeryHideAsync(testFile, "Workflow");
        var check3 = await _sheetCommands.GetVisibilityAsync(testFile, "Workflow");
        Assert.Equal(SheetVisibility.VeryHidden, check3.Visibility);

        // Show it again
        await _sheetCommands.ShowAsync(testFile, "Workflow");
        var check4 = await _sheetCommands.GetVisibilityAsync(testFile, "Workflow");
        Assert.Equal(SheetVisibility.Visible, check4.Visibility);

        // Save changes
    }
}
