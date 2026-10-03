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
        Assert.True(_commands.Hide(batch).Success);
        var before = _commands.GetInfo(batch);
        Assert.True(before.Success, before.ErrorMessage);
        Assert.False(before.IsVisible);

        // Act
        var result = _commands.Show(batch);

        // Assert
        Assert.True(result.Success, $"Show failed: {result.ErrorMessage}");
        Assert.Equal("show", result.Action);

        // Verify via GetInfo
        var info = _commands.GetInfo(batch);
        Assert.True(info.Success, info.ErrorMessage);
        Assert.True(info.IsVisible);
        Assert.True(info.IsForeground);

        // Cleanup: hide again so tests don't leave visible Excel windows
        RequireSuccess(_commands.Hide(batch));
    }

    [Fact]
    public void Hide_MakesExcelInvisible()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Ensure visible first
        Assert.True(_commands.Show(batch).Success);
        var before = _commands.GetInfo(batch);
        Assert.True(before.Success, before.ErrorMessage);
        Assert.True(before.IsVisible);

        // Act
        var result = _commands.Hide(batch);

        // Assert
        Assert.True(result.Success, $"Hide failed: {result.ErrorMessage}");
        Assert.Equal("hide", result.Action);

        // Verify via GetInfo
        var info = _commands.GetInfo(batch);
        Assert.True(info.Success, info.ErrorMessage);
        Assert.False(info.IsVisible);
    }

    [Fact]
    public void BringToFront_WhenHidden_ReturnsGuidanceMessage()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        Assert.True(_commands.Hide(batch).Success);

        // Act
        var result = _commands.BringToFront(batch);

        // Assert - should succeed but with guidance message
        Assert.True(result.Success);
        Assert.Equal("bring-to-front", result.Action);
        Assert.Contains("show", result.Message, StringComparison.OrdinalIgnoreCase);
        var after = _commands.GetInfo(batch);
        Assert.True(after.Success, after.ErrorMessage);
        Assert.False(after.IsVisible);
        Assert.False(after.IsForeground);
    }

    [Theory]
    [InlineData("normal")]
    [InlineData("maximized")]
    public void BringToFront_WhenVisible_Succeeds(string originalState)
    {
        // Arrange
        var batch = _fixture.BatchToken;
        RequireSuccess(_commands.Show(batch));
        RequireSuccess(_commands.SetState(batch, originalState));
        RequireSuccess(_commands.SetState(batch, "minimized"));
        var before = RequireSuccess(_commands.GetInfo(batch));
        Assert.True(before.IsVisible);
        Assert.Equal("minimized", before.WindowState);
        Assert.False(before.IsForeground);

        // Act
        var result = _commands.BringToFront(batch);

        // Assert
        Assert.True(result.Success);
        Assert.Equal("bring-to-front", result.Action);
        Assert.Contains("foreground", result.Message, StringComparison.OrdinalIgnoreCase);
        var after = RequireSuccess(_commands.GetInfo(batch));
        Assert.True(after.IsVisible);
        Assert.True(after.IsForeground);
        Assert.Equal(originalState, after.WindowState);

        // Cleanup
        RequireSuccess(_commands.Hide(batch));
    }

    [Theory]
    [InlineData("normal")]
    [InlineData("maximized")]
    public void Show_WhenHiddenAndMinimized_RestoresWindowAndForeground(string originalState)
    {
        var batch = _fixture.BatchToken;
        RequireSuccess(_commands.SetState(batch, originalState));
        RequireSuccess(_commands.SetState(batch, "minimized"));
        RequireSuccess(_commands.Hide(batch));
        Assert.False(RequireSuccess(_commands.GetInfo(batch)).IsVisible);

        RequireSuccess(_commands.Show(batch));

        var after = RequireSuccess(_commands.GetInfo(batch));
        Assert.True(after.IsVisible);
        Assert.True(after.IsForeground);
        Assert.Equal(originalState, after.WindowState);
    }

    [Theory]
    [InlineData("normal")]
    [InlineData("maximized")]
    public void RestoreMinimizedWindow_NativeControl_PreservesOriginalState(string originalState)
    {
        var batch = _fixture.BatchToken;
        RequireSuccess(_commands.SetState(batch, originalState));
        RequireSuccess(_commands.SetState(batch, "minimized"));
        var before = RequireSuccess(_commands.GetInfo(batch));
        Assert.Equal("minimized", before.WindowState);
        Assert.False(before.IsForeground);

        _fixture.ExecuteRawVerification((context, cancellationToken) =>
        {
            var hwnd = new IntPtr(context.App.Hwnd);
            _ = ShowWindow(hwnd, 9);
            _ = SetForegroundWindow(hwnd);
        });

        var after = RequireSuccess(_commands.GetInfo(batch));
        Assert.True(after.IsVisible);
        Assert.True(after.IsForeground);
        Assert.Equal(originalState, after.WindowState);
    }
}
