// <copyright file="WindowCommandsTests.State.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for SetState and SetPosition operations.
/// </summary>
public sealed partial class PersistentServiceWindowTests
{
    [Theory]
    [InlineData("normal")]
    [InlineData("maximized")]
    [InlineData("minimized")]
    public void SetState_ValidStates_Succeed(string state)
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var baseline = state == "normal" ? "maximized" : "normal";
        RequireSuccess(_commands.SetState(batch, baseline));
        Assert.Equal(baseline, RequireSuccess(_commands.GetInfo(batch)).WindowState);
        RequireSuccess(_commands.Hide(batch));

        // Act
        var result = _commands.SetState(batch, state);

        // Assert
        Assert.True(result.Success, $"SetState '{state}' failed: {result.ErrorMessage}");
        Assert.Equal("set-state", result.Action);
        Assert.Contains(state, result.Message, StringComparison.OrdinalIgnoreCase);
        var actual = _commands.GetInfo(batch);
        Assert.True(actual.Success, actual.ErrorMessage);
        Assert.Equal(state, actual.WindowState);
        Assert.True(actual.IsVisible);

        // Cleanup
        RequireSuccess(_commands.Hide(batch));
    }

    [Fact]
    public void SetState_InvalidState_Throws()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        Assert.True(_commands.Hide(batch).Success);
        var before = _commands.GetInfo(batch);
        Assert.True(before.Success, before.ErrorMessage);

        // Act & Assert
        var exception = Assert.Throws<ArgumentException>(() => _commands.SetState(batch, "invalid-state"));
        Assert.Contains("Unknown window state", exception.Message, StringComparison.Ordinal);
        var after = _commands.GetInfo(batch);
        Assert.True(after.Success, after.ErrorMessage);
        Assert.Equal(before.IsVisible, after.IsVisible);
        Assert.Equal(before.WindowState, after.WindowState);
    }

    [Fact]
    public void SetPosition_AllParameters_UpdatesPosition()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        (double Left, double Top, double Width, double Height) expected = default;
        _fixture.ExecuteRawVerification((context, _) =>
        {
            context.App.Visible = true;
            context.App.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlNormal;
            context.App.Left = 100;
            context.App.Top = 50;
            context.App.Width = 800;
            context.App.Height = 600;
            // Excel rounds window coordinates and may adjust them while resizing.
            expected = (context.App.Left, context.App.Top, context.App.Width, context.App.Height);
            context.App.Left = 20;
            context.App.Top = 20;
            context.App.Width = 400;
            context.App.Height = 300;
        });
        Assert.True(_commands.Hide(batch).Success);

        // Act
        var result = _commands.SetPosition(batch, left: 100, top: 50, width: 800, height: 600);

        // Assert
        Assert.True(result.Success, $"SetPosition failed: {result.ErrorMessage}");
        Assert.Equal("set-position", result.Action);

        // Verify position via GetInfo
        var info = _commands.GetInfo(batch);
        Assert.True(info.Success, info.ErrorMessage);
        Assert.True(info.IsVisible, "SetPosition should make Excel visible");
        Assert.Equal(expected.Left, info.Left, 1d);
        Assert.Equal(expected.Top, info.Top, 1d);
        Assert.Equal(expected.Width, info.Width, 1d);
        Assert.Equal(expected.Height, info.Height, 1d);

        // Cleanup
        RequireSuccess(_commands.Hide(batch));
    }

    [Fact]
    public void SetPosition_PartialParameters_OnlyUpdatesProvided()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        RequireSuccess(_commands.Show(batch));
        RequireSuccess(_commands.SetState(batch, "normal"));
        RequireSuccess(_commands.SetPosition(batch, left: 20));
        var beforeInfo = _commands.GetInfo(batch);
        Assert.True(beforeInfo.Success, beforeInfo.ErrorMessage);

        // Act - only change left position
        var result = _commands.SetPosition(batch, left: 200);

        // Assert
        Assert.True(result.Success);

        var afterInfo = _commands.GetInfo(batch);
        Assert.True(afterInfo.Success, afterInfo.ErrorMessage);
        Assert.Equal(200, afterInfo.Left, 1.0); // Allow small floating-point tolerance
        Assert.Equal(beforeInfo.Top, afterInfo.Top, 1d);
        Assert.Equal(beforeInfo.Width, afterInfo.Width, 1d);
        Assert.Equal(beforeInfo.Height, afterInfo.Height, 1d);

        // Cleanup
        RequireSuccess(_commands.Hide(batch));
    }

    [Fact]
    public void SetState_MakesHiddenWindowVisible()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        RequireSuccess(_commands.SetState(batch, "maximized"));
        Assert.Equal("maximized", RequireSuccess(_commands.GetInfo(batch)).WindowState);
        RequireSuccess(_commands.Hide(batch));

        // Verify hidden
        var beforeInfo = _commands.GetInfo(batch);
        Assert.True(beforeInfo.Success, beforeInfo.ErrorMessage);
        Assert.False(beforeInfo.IsVisible);

        // Act
        var result = _commands.SetState(batch, "normal");

        // Assert - should auto-show
        Assert.True(result.Success);
        var afterInfo = _commands.GetInfo(batch);
        Assert.True(afterInfo.Success, afterInfo.ErrorMessage);
        Assert.True(afterInfo.IsVisible);
        Assert.Equal("normal", afterInfo.WindowState);

        // Cleanup
        RequireSuccess(_commands.Hide(batch));
    }
}
