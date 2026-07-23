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
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = $"Hide_{Guid.NewGuid():N}"[..31];
        _sheetCommands.Create(batch, sheetName);

        // Act
        _sheetCommands.SetVisibility(batch, sheetName, SheetVisibility.Hidden);  // SetVisibility throws on error

        // Assert - reaching here means set succeeded

        // Verify by reading visibility
        var getResult = _sheetCommands.GetVisibility(batch, sheetName);
        Assert.True(getResult.Success);
        Assert.Equal(SheetVisibility.Hidden, getResult.Visibility);
        Assert.Equal("Hidden", getResult.VisibilityName);
    }
    /// <inheritdoc/>

    [Fact]
    public void SetVisibility_ToVeryHidden_WorksCorrectly()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = $"VHide_{Guid.NewGuid():N}"[..31];
        _sheetCommands.Create(batch, sheetName);

        // Act
        _sheetCommands.SetVisibility(batch, sheetName, SheetVisibility.VeryHidden);  // SetVisibility throws on error

        // Assert - reaching here means set succeeded

        var getResult = _sheetCommands.GetVisibility(batch, sheetName);
        Assert.True(getResult.Success);
        Assert.Equal(SheetVisibility.VeryHidden, getResult.Visibility);
        Assert.Equal("VeryHidden", getResult.VisibilityName);
    }
    /// <inheritdoc/>

    [Fact]
    public void Show_HiddenSheet_MakesVisible()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = $"ShowH_{Guid.NewGuid():N}"[..31];
        _sheetCommands.Create(batch, sheetName);
        _sheetCommands.Hide(batch, sheetName);

        // Verify it's hidden
        var hiddenCheck = _sheetCommands.GetVisibility(batch, sheetName);
        Assert.Equal(SheetVisibility.Hidden, hiddenCheck.Visibility);

        // Act - Show the sheet
        _sheetCommands.Show(batch, sheetName);  // Show throws on error

        // Assert - reaching here means show succeeded

        var visibleCheck = _sheetCommands.GetVisibility(batch, sheetName);
        Assert.Equal(SheetVisibility.Visible, visibleCheck.Visibility);
    }
    /// <inheritdoc/>

    [Fact]
    public void Show_VeryHiddenSheet_MakesVisible()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = $"ShowVH_{Guid.NewGuid():N}"[..31];
        _sheetCommands.Create(batch, sheetName);
        _sheetCommands.VeryHide(batch, sheetName);

        // Verify it's very hidden
        var veryHiddenCheck = _sheetCommands.GetVisibility(batch, sheetName);
        Assert.Equal(SheetVisibility.VeryHidden, veryHiddenCheck.Visibility);

        // Act - Show the sheet
        _sheetCommands.Show(batch, sheetName);  // Show throws on error

        // Assert - reaching here means show succeeded

        var visibleCheck = _sheetCommands.GetVisibility(batch, sheetName);
        Assert.Equal(SheetVisibility.Visible, visibleCheck.Visibility);
    }
    /// <inheritdoc/>

    [Fact]
    public void Hide_VisibleSheet_MakesHidden()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = $"HideMe_{Guid.NewGuid():N}"[..31];
        _sheetCommands.Create(batch, sheetName);

        // Act
        _sheetCommands.Hide(batch, sheetName);  // Hide throws on error

        // Assert - reaching here means hide succeeded

        var getResult = _sheetCommands.GetVisibility(batch, sheetName);
        Assert.Equal(SheetVisibility.Hidden, getResult.Visibility);
    }
    /// <inheritdoc/>

    [Fact]
    public void VeryHide_VeryHidesVisibleSheet()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = $"VHide_{Guid.NewGuid():N}"[..31];
        _sheetCommands.Create(batch, sheetName);

        // Act
        _sheetCommands.VeryHide(batch, sheetName);  // VeryHide throws on error

        // Assert - reaching here means veryhide succeeded

        var getResult = _sheetCommands.GetVisibility(batch, sheetName);
        Assert.Equal(SheetVisibility.VeryHidden, getResult.Visibility);
    }
    /// <inheritdoc/>

    [Fact]
    public void GetVisibility_ForVisibleSheet_ReturnsVisible()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = $"Vis_{Guid.NewGuid():N}"[..31];
        _sheetCommands.Create(batch, sheetName);

        // Act
        var result = _sheetCommands.GetVisibility(batch, sheetName);

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
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);

        // Act & Assert - Should throw InvalidOperationException when sheet not found
        var exception = Assert.Throws<InvalidOperationException>(
            () => _sheetCommands.SetVisibility(batch, $"NonExist_{Guid.NewGuid():N}", SheetVisibility.Hidden));
        Assert.Contains("not found", exception.Message);
    }
    /// <inheritdoc/>

    [Fact]
    public void Visibility_CompleteWorkflow_AllLevelsWork()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = $"WFlow_{Guid.NewGuid():N}"[..31];
        _sheetCommands.Create(batch, sheetName);

        // Act & Assert - Test complete visibility workflow

        // Start visible
        var check1 = _sheetCommands.GetVisibility(batch, sheetName);
        Assert.Equal(SheetVisibility.Visible, check1.Visibility);

        // Hide it
        _sheetCommands.Hide(batch, sheetName);
        var check2 = _sheetCommands.GetVisibility(batch, sheetName);
        Assert.Equal(SheetVisibility.Hidden, check2.Visibility);

        // Very hide it
        _sheetCommands.VeryHide(batch, sheetName);
        var check3 = _sheetCommands.GetVisibility(batch, sheetName);
        Assert.Equal(SheetVisibility.VeryHidden, check3.Visibility);

        // Show it again
        _sheetCommands.Show(batch, sheetName);
        var check4 = _sheetCommands.GetVisibility(batch, sheetName);
        Assert.Equal(SheetVisibility.Visible, check4.Visibility);
    }
}




