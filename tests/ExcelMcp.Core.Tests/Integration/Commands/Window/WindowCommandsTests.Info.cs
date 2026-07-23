// <copyright file="WindowCommandsTests.Info.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.Window;

/// <summary>
/// Tests for GetInfo operation.
/// </summary>
public partial class WindowCommandsTests
{
    [Fact]
    public void GetInfo_WhenHidden_ReturnsHiddenState()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        _commands.Hide(batch);

        // Act
        var info = _commands.GetInfo(batch);

        // Assert
        Assert.True(info.Success, $"GetInfo failed: {info.ErrorMessage}");
        Assert.Equal("get-info", info.Action);
        Assert.False(info.IsVisible);
        Assert.Contains("hidden", info.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetInfo_WhenVisible_ReturnsPositionAndSize()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        _commands.Show(batch);

        // Act
        var info = _commands.GetInfo(batch);

        // Assert
        Assert.True(info.Success);
        Assert.True(info.IsVisible);
        Assert.NotEmpty(info.WindowState);
        Assert.True(info.Width > 0, "Width should be positive when visible");
        Assert.True(info.Height > 0, "Height should be positive when visible");

        // Cleanup
        _commands.Hide(batch);
    }

    [Fact]
    public void GetInfo_WhenMaximized_ReportsMaximizedState()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        _commands.Show(batch);
        _commands.SetState(batch, "maximized");

        // Act
        var info = _commands.GetInfo(batch);

        // Assert
        Assert.True(info.Success);
        Assert.Equal("maximized", info.WindowState);

        // Cleanup
        _commands.Hide(batch);
    }
}
