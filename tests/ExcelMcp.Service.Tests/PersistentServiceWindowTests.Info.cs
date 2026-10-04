// <copyright file="WindowCommandsTests.Info.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for GetInfo operation.
/// </summary>
public sealed partial class PersistentServiceWindowTests
{
    [Fact]
    public void GetInfo_WhenHidden_ReturnsHiddenState()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        Assert.True(_commands.Hide(batch).Success);

        // Act
        var info = _commands.GetInfo(batch);

        // Assert
        Assert.True(info.Success, $"GetInfo failed: {info.ErrorMessage}");
        Assert.Equal("get-info", info.Action);
        Assert.False(info.IsVisible);
        Assert.False(info.IsForeground);
        Assert.Contains("hidden", info.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetInfo_WhenVisible_ReturnsPositionAndSize()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var expected = _fixture.ExecuteRawVerification((context, _) =>
        {
            context.App.Visible = true;
            context.App.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlNormal;
            context.App.Left = 97;
            context.App.Top = 43;
            context.App.Width = 711;
            context.App.Height = 523;
            return (context.App.Left, context.App.Top, context.App.Width, context.App.Height);
        });

        // Act
        var info = _commands.GetInfo(batch);

        // Assert
        Assert.True(info.Success);
        Assert.True(info.IsVisible);
        Assert.Equal("normal", info.WindowState);
        Assert.Equal(expected.Left, info.Left);
        Assert.Equal(expected.Top, info.Top);
        Assert.Equal(expected.Width, info.Width);
        Assert.Equal(expected.Height, info.Height);

        // Cleanup
        RequireSuccess(_commands.Hide(batch));
    }

    [Fact]
    public void GetInfo_WhenMaximized_ReportsMaximizedState()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        Assert.True(_commands.Show(batch).Success);
        Assert.True(_commands.SetState(batch, "maximized").Success);

        // Act
        var info = _commands.GetInfo(batch);

        // Assert
        Assert.True(info.Success);
        Assert.Equal("maximized", info.WindowState);

        // Cleanup
        RequireSuccess(_commands.Hide(batch));
    }
}
