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
    public void SetVisibility_ToHidden_WorksCorrectly()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests),
            nameof(SetVisibility_ToHidden_WorksCorrectly),
            _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "HideTest");

        // Act
        var setResult = _sheetCommands.SetVisibility(batch, "HideTest", SheetVisibility.Hidden);

        // Assert
        Assert.True(setResult.Success, $"SetVisibility failed: {setResult.ErrorMessage}");

        // Verify by reading visibility
        var getResult = _sheetCommands.GetVisibility(batch, "HideTest");
        Assert.True(getResult.Success);
        Assert.Equal(SheetVisibility.Hidden, getResult.Visibility);
        Assert.Equal("Hidden", getResult.VisibilityName);

        // Save changes
    }
    /// <inheritdoc/>

    [Fact]
    public void SetVisibility_ToVeryHidden_WorksCorrectly()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests),
            nameof(SetVisibility_ToVeryHidden_WorksCorrectly),
            _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "VeryHideTest");

        // Act
        var setResult = _sheetCommands.SetVisibility(batch, "VeryHideTest", SheetVisibility.VeryHidden);

        // Assert
        Assert.True(setResult.Success);

        var getResult = _sheetCommands.GetVisibility(batch, "VeryHideTest");
        Assert.True(getResult.Success);
        Assert.Equal(SheetVisibility.VeryHidden, getResult.Visibility);
        Assert.Equal("VeryHidden", getResult.VisibilityName);

        // Save changes
    }
    /// <inheritdoc/>

    [Fact]
    public void Show_HiddenSheet_MakesVisible()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests),
            nameof(Show_HiddenSheet_MakesVisible),
            _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "ShowTest");
        _sheetCommands.Hide(batch, "ShowTest");

        // Verify it's hidden
        var hiddenCheck = _sheetCommands.GetVisibility(batch, "ShowTest");
        Assert.Equal(SheetVisibility.Hidden, hiddenCheck.Visibility);

        // Act - Show the sheet
        var showResult = _sheetCommands.Show(batch, "ShowTest");

        // Assert
        Assert.True(showResult.Success);

        var visibleCheck = _sheetCommands.GetVisibility(batch, "ShowTest");
        Assert.Equal(SheetVisibility.Visible, visibleCheck.Visibility);

        // Save changes
    }
    /// <inheritdoc/>

    [Fact]
    public void Show_VeryHiddenSheet_MakesVisible()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests),
            nameof(Show_VeryHiddenSheet_MakesVisible),
            _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "VeryHideShowTest");
        _sheetCommands.VeryHide(batch, "VeryHideShowTest");

        // Verify it's very hidden
        var veryHiddenCheck = _sheetCommands.GetVisibility(batch, "VeryHideShowTest");
        Assert.Equal(SheetVisibility.VeryHidden, veryHiddenCheck.Visibility);

        // Act - Show the sheet
        var showResult = _sheetCommands.Show(batch, "VeryHideShowTest");

        // Assert
        Assert.True(showResult.Success);

        var visibleCheck = _sheetCommands.GetVisibility(batch, "VeryHideShowTest");
        Assert.Equal(SheetVisibility.Visible, visibleCheck.Visibility);

        // Save changes
    }
    /// <inheritdoc/>

    [Fact]
    public void Hide_VisibleSheet_MakesHidden()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests),
            nameof(Hide_VisibleSheet_MakesHidden),
            _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "HideMe");

        // Act
        var hideResult = _sheetCommands.Hide(batch, "HideMe");

        // Assert
        Assert.True(hideResult.Success);

        var getResult = _sheetCommands.GetVisibility(batch, "HideMe");
        Assert.Equal(SheetVisibility.Hidden, getResult.Visibility);

        // Save changes
    }
    /// <inheritdoc/>

    [Fact]
    public void VeryHide_VeryHidesVisibleSheet()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests),
            nameof(VeryHide_VeryHidesVisibleSheet),
            _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "VeryHideMe");

        // Act
        var veryHideResult = _sheetCommands.VeryHide(batch, "VeryHideMe");

        // Assert
        Assert.True(veryHideResult.Success);

        var getResult = _sheetCommands.GetVisibility(batch, "VeryHideMe");
        Assert.Equal(SheetVisibility.VeryHidden, getResult.Visibility);

        // Save changes
    }
    /// <inheritdoc/>

    [Fact]
    public void GetVisibility_ForVisibleSheet_ReturnsVisible()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests),
            nameof(GetVisibility_ForVisibleSheet_ReturnsVisible),
            _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "VisibleSheet");

        // Act
        var result = _sheetCommands.GetVisibility(batch, "VisibleSheet");

        // Assert
        Assert.True(result.Success);
        Assert.Equal(SheetVisibility.Visible, result.Visibility);
        Assert.Equal("Visible", result.VisibilityName);
    }
    /// <inheritdoc/>

    [Fact]
    public void SetVisibility_WithNonExistentSheet_ThrowsException()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests),
            nameof(SetVisibility_WithNonExistentSheet_ThrowsException),
            _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act & Assert - Should throw InvalidOperationException when sheet not found
        var exception = Assert.Throws<InvalidOperationException>(
            () => _sheetCommands.SetVisibility(batch, "NonExistent", SheetVisibility.Hidden));
        Assert.Contains("not found", exception.Message);
    }
    /// <inheritdoc/>

    [Fact]
    public void Visibility_CompleteWorkflow_AllLevelsWork()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests),
            nameof(Visibility_CompleteWorkflow_AllLevelsWork),
            _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "Workflow");

        // Act & Assert - Test complete visibility workflow

        // Start visible
        var check1 = _sheetCommands.GetVisibility(batch, "Workflow");
        Assert.Equal(SheetVisibility.Visible, check1.Visibility);

        // Hide it
        _sheetCommands.Hide(batch, "Workflow");
        var check2 = _sheetCommands.GetVisibility(batch, "Workflow");
        Assert.Equal(SheetVisibility.Hidden, check2.Visibility);

        // Very hide it
        _sheetCommands.VeryHide(batch, "Workflow");
        var check3 = _sheetCommands.GetVisibility(batch, "Workflow");
        Assert.Equal(SheetVisibility.VeryHidden, check3.Visibility);

        // Show it again
        _sheetCommands.Show(batch, "Workflow");
        var check4 = _sheetCommands.GetVisibility(batch, "Workflow");
        Assert.Equal(SheetVisibility.Visible, check4.Visibility);

        // Save changes
    }
}
