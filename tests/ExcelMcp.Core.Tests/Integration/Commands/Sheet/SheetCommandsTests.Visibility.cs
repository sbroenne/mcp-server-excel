using Sbroenne.ExcelMcp.ComInterop.Session;
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

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        await _sheetCommands.CreateAsync(batch, "HideTest");

        // Act
        var setResult = await _sheetCommands.SetVisibilityAsync(batch, "HideTest", SheetVisibility.Hidden);

        // Assert
        Assert.True(setResult.Success, $"SetVisibility failed: {setResult.ErrorMessage}");

        // Verify by reading visibility
        var getResult = await _sheetCommands.GetVisibilityAsync(batch, "HideTest");
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

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        await _sheetCommands.CreateAsync(batch, "VeryHideTest");

        // Act
        var setResult = await _sheetCommands.SetVisibilityAsync(batch, "VeryHideTest", SheetVisibility.VeryHidden);

        // Assert
        Assert.True(setResult.Success);

        var getResult = await _sheetCommands.GetVisibilityAsync(batch, "VeryHideTest");
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

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        await _sheetCommands.CreateAsync(batch, "ShowTest");
        await _sheetCommands.HideAsync(batch, "ShowTest");

        // Verify it's hidden
        var hiddenCheck = await _sheetCommands.GetVisibilityAsync(batch, "ShowTest");
        Assert.Equal(SheetVisibility.Hidden, hiddenCheck.Visibility);

        // Act - Show the sheet
        var showResult = await _sheetCommands.ShowAsync(batch, "ShowTest");

        // Assert
        Assert.True(showResult.Success);

        var visibleCheck = await _sheetCommands.GetVisibilityAsync(batch, "ShowTest");
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

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        await _sheetCommands.CreateAsync(batch, "VeryHideShowTest");
        await _sheetCommands.VeryHideAsync(batch, "VeryHideShowTest");

        // Verify it's very hidden
        var veryHiddenCheck = await _sheetCommands.GetVisibilityAsync(batch, "VeryHideShowTest");
        Assert.Equal(SheetVisibility.VeryHidden, veryHiddenCheck.Visibility);

        // Act - Show the sheet
        var showResult = await _sheetCommands.ShowAsync(batch, "VeryHideShowTest");

        // Assert
        Assert.True(showResult.Success);

        var visibleCheck = await _sheetCommands.GetVisibilityAsync(batch, "VeryHideShowTest");
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

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        await _sheetCommands.CreateAsync(batch, "HideMe");

        // Act
        var hideResult = await _sheetCommands.HideAsync(batch, "HideMe");

        // Assert
        Assert.True(hideResult.Success);

        var getResult = await _sheetCommands.GetVisibilityAsync(batch, "HideMe");
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

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        await _sheetCommands.CreateAsync(batch, "VeryHideMe");

        // Act
        var veryHideResult = await _sheetCommands.VeryHideAsync(batch, "VeryHideMe");

        // Assert
        Assert.True(veryHideResult.Success);

        var getResult = await _sheetCommands.GetVisibilityAsync(batch, "VeryHideMe");
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

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        await _sheetCommands.CreateAsync(batch, "VisibleSheet");

        // Act
        var result = await _sheetCommands.GetVisibilityAsync(batch, "VisibleSheet");

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

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act
        var result = await _sheetCommands.SetVisibilityAsync(batch, "NonExistent", SheetVisibility.Hidden);

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

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        await _sheetCommands.CreateAsync(batch, "Workflow");

        // Act & Assert - Test complete visibility workflow

        // Start visible
        var check1 = await _sheetCommands.GetVisibilityAsync(batch, "Workflow");
        Assert.Equal(SheetVisibility.Visible, check1.Visibility);

        // Hide it
        await _sheetCommands.HideAsync(batch, "Workflow");
        var check2 = await _sheetCommands.GetVisibilityAsync(batch, "Workflow");
        Assert.Equal(SheetVisibility.Hidden, check2.Visibility);

        // Very hide it
        await _sheetCommands.VeryHideAsync(batch, "Workflow");
        var check3 = await _sheetCommands.GetVisibilityAsync(batch, "Workflow");
        Assert.Equal(SheetVisibility.VeryHidden, check3.Visibility);

        // Show it again
        await _sheetCommands.ShowAsync(batch, "Workflow");
        var check4 = await _sheetCommands.GetVisibilityAsync(batch, "Workflow");
        Assert.Equal(SheetVisibility.Visible, check4.Visibility);

        // Save changes
    }
}
