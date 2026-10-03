using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for worksheet visibility operations
/// </summary>
public sealed partial class PersistentServiceSheetTests
{
    /// <inheritdoc/>

    [Fact]
    public void SetVisibility_ToHidden_WorksCorrectly()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = $"Hide_{Guid.NewGuid():N}"[..31];
        _fixture.CreateNamedTestSheet(batch, sheetName);

        // Act
        _sheetCommands.SetVisibility(batch, sheetName, SheetVisibility.Hidden);  // SetVisibility throws on error

        // Assert - reaching here means set succeeded

        // Verify by reading visibility
        var getResult = _sheetCommands.GetVisibility(batch, sheetName);
        RequireSuccess(getResult);
        Assert.Equal(SheetVisibility.Hidden, getResult.Visibility);
        Assert.Equal("Hidden", getResult.VisibilityName);
    }
    /// <inheritdoc/>

    [Fact]
    public void SetVisibility_ToVeryHidden_WorksCorrectly()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = $"VHide_{Guid.NewGuid():N}"[..31];
        _fixture.CreateNamedTestSheet(batch, sheetName);

        // Act
        _sheetCommands.SetVisibility(batch, sheetName, SheetVisibility.VeryHidden);  // SetVisibility throws on error

        // Assert - reaching here means set succeeded

        var getResult = _sheetCommands.GetVisibility(batch, sheetName);
        RequireSuccess(getResult);
        Assert.Equal(SheetVisibility.VeryHidden, getResult.Visibility);
        Assert.Equal("VeryHidden", getResult.VisibilityName);
    }
    /// <inheritdoc/>

    [Fact]
    public void Show_HiddenSheet_MakesVisible()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = $"ShowH_{Guid.NewGuid():N}"[..31];
        _fixture.CreateNamedTestSheet(batch, sheetName);
        _sheetCommands.Hide(batch, sheetName);

        // Verify it's hidden
        var hiddenCheck = _sheetCommands.GetVisibility(batch, sheetName);
        RequireSuccess(hiddenCheck);
        Assert.Equal(SheetVisibility.Hidden, hiddenCheck.Visibility);

        // Act - Show the sheet
        _sheetCommands.Show(batch, sheetName);  // Show throws on error

        // Assert - reaching here means show succeeded

        var visibleCheck = _sheetCommands.GetVisibility(batch, sheetName);
        RequireSuccess(visibleCheck);
        Assert.Equal(SheetVisibility.Visible, visibleCheck.Visibility);
    }
    /// <inheritdoc/>

    [Fact]
    public void Show_VeryHiddenSheet_MakesVisible()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = $"ShowVH_{Guid.NewGuid():N}"[..31];
        _fixture.CreateNamedTestSheet(batch, sheetName);
        _sheetCommands.VeryHide(batch, sheetName);

        // Verify it's very hidden
        var veryHiddenCheck = _sheetCommands.GetVisibility(batch, sheetName);
        RequireSuccess(veryHiddenCheck);
        Assert.Equal(SheetVisibility.VeryHidden, veryHiddenCheck.Visibility);

        // Act - Show the sheet
        _sheetCommands.Show(batch, sheetName);  // Show throws on error

        // Assert - reaching here means show succeeded

        var visibleCheck = _sheetCommands.GetVisibility(batch, sheetName);
        RequireSuccess(visibleCheck);
        Assert.Equal(SheetVisibility.Visible, visibleCheck.Visibility);
    }
    /// <inheritdoc/>

    [Fact]
    public void Hide_VisibleSheet_MakesHidden()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = $"HideMe_{Guid.NewGuid():N}"[..31];
        _fixture.CreateNamedTestSheet(batch, sheetName);

        // Act
        _sheetCommands.Hide(batch, sheetName);  // Hide throws on error

        // Assert - reaching here means hide succeeded

        var getResult = _sheetCommands.GetVisibility(batch, sheetName);
        RequireSuccess(getResult);
        Assert.Equal(SheetVisibility.Hidden, getResult.Visibility);
    }
    /// <inheritdoc/>

    [Fact]
    public void VeryHide_VeryHidesVisibleSheet()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = $"VHide_{Guid.NewGuid():N}"[..31];
        _fixture.CreateNamedTestSheet(batch, sheetName);

        // Act
        _sheetCommands.VeryHide(batch, sheetName);  // VeryHide throws on error

        // Assert - reaching here means veryhide succeeded

        var getResult = _sheetCommands.GetVisibility(batch, sheetName);
        RequireSuccess(getResult);
        Assert.Equal(SheetVisibility.VeryHidden, getResult.Visibility);
    }
    /// <inheritdoc/>

    [Fact]
    public void GetVisibility_ForVisibleSheet_ReturnsVisible()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = $"Vis_{Guid.NewGuid():N}"[..31];
        _fixture.CreateNamedTestSheet(batch, sheetName);

        // Act
        var result = _sheetCommands.GetVisibility(batch, sheetName);

        // Assert
        RequireSuccess(result);
        Assert.Equal(SheetVisibility.Visible, result.Visibility);
        Assert.Equal("Visible", result.VisibilityName);
    }
    /// <inheritdoc/>

    [Fact]
    public void SetVisibility_WithNonExistentSheet_ThrowsException()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        _sheetCommands.VeryHide(batch, sheetName);
        var before = _sheetCommands.GetVisibility(batch, sheetName);
        RequireSuccess(before);
        Assert.Equal(SheetVisibility.VeryHidden, before.Visibility);

        // Act & Assert - Should throw InvalidOperationException when sheet not found
        var exception = Assert.Throws<InvalidOperationException>(
            () => _sheetCommands.SetVisibility(batch, $"NonExist_{Guid.NewGuid():N}", SheetVisibility.Hidden));
        Assert.Contains("not found", exception.Message);
        var after = _sheetCommands.GetVisibility(batch, sheetName);
        RequireSuccess(after);
        Assert.Equal(before.Visibility, after.Visibility);
        Assert.Equal(before.VisibilityName, after.VisibilityName);
    }
    /// <inheritdoc/>

    [Fact]
    public void Visibility_CompleteWorkflow_AllLevelsWork()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = $"WFlow_{Guid.NewGuid():N}"[..31];
        _fixture.CreateNamedTestSheet(batch, sheetName);

        // Act & Assert - Test complete visibility workflow

        // Start visible
        var check1 = _sheetCommands.GetVisibility(batch, sheetName);
        RequireSuccess(check1);
        Assert.Equal(SheetVisibility.Visible, check1.Visibility);

        // Hide it
        _sheetCommands.Hide(batch, sheetName);
        var check2 = _sheetCommands.GetVisibility(batch, sheetName);
        RequireSuccess(check2);
        Assert.Equal(SheetVisibility.Hidden, check2.Visibility);

        // Very hide it
        _sheetCommands.VeryHide(batch, sheetName);
        var check3 = _sheetCommands.GetVisibility(batch, sheetName);
        RequireSuccess(check3);
        Assert.Equal(SheetVisibility.VeryHidden, check3.Visibility);

        // Show it again
        _sheetCommands.Show(batch, sheetName);
        var check4 = _sheetCommands.GetVisibility(batch, sheetName);
        RequireSuccess(check4);
        Assert.Equal(SheetVisibility.Visible, check4.Visibility);
    }
}



