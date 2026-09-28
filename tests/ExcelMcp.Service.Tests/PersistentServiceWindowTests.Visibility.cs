// <copyright file="WindowCommandsTests.Visibility.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for Show, Hide, and BringToFront operations.
/// </summary>
public sealed partial class PersistentServiceWindowTests
{
    [Fact]
    public void Show_MakesExcelVisible()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Start hidden
        _commands.Hide(batch);

        // Act
        var result = _commands.Show(batch);

        // Assert
        Assert.True(result.Success, $"Show failed: {result.ErrorMessage}");
        Assert.Equal("show", result.Action);

        // Verify via GetInfo
        var info = _commands.GetInfo(batch);
        Assert.True(info.IsVisible);

        // Cleanup: hide again so tests don't leave visible Excel windows
        _commands.Hide(batch);
    }

    [Fact]
    public void Hide_MakesExcelInvisible()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Ensure visible first
        _commands.Show(batch);

        // Act
        var result = _commands.Hide(batch);

        // Assert
        Assert.True(result.Success, $"Hide failed: {result.ErrorMessage}");
        Assert.Equal("hide", result.Action);

        // Verify via GetInfo
        var info = _commands.GetInfo(batch);
        Assert.False(info.IsVisible);
    }

    [Fact]
    public void Show_Then_Hide_Roundtrip()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Act & Assert - Show
        var showResult = _commands.Show(batch);
        Assert.True(showResult.Success);

        var infoAfterShow = _commands.GetInfo(batch);
        Assert.True(infoAfterShow.IsVisible);

        // Act & Assert - Hide
        var hideResult = _commands.Hide(batch);
        Assert.True(hideResult.Success);

        var infoAfterHide = _commands.GetInfo(batch);
        Assert.False(infoAfterHide.IsVisible);
    }

    [Fact]
    public void BringToFront_WhenHidden_ReturnsGuidanceMessage()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        _commands.Hide(batch);

        // Act
        var result = _commands.BringToFront(batch);

        // Assert - should succeed but with guidance message
        Assert.True(result.Success);
        Assert.Equal("bring-to-front", result.Action);
        Assert.Contains("show", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void BringToFront_WhenVisible_Succeeds()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        _commands.Show(batch);

        // Act
        var result = _commands.BringToFront(batch);

        // Assert
        Assert.True(result.Success);
        Assert.Equal("bring-to-front", result.Action);
        Assert.Contains("foreground", result.Message, StringComparison.OrdinalIgnoreCase);

        // Cleanup
        _commands.Hide(batch);
    }
}
