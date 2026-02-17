// <copyright file="WindowCommandsTests.Arrange.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.Window;

/// <summary>
/// Tests for Arrange preset operations.
/// </summary>
public partial class WindowCommandsTests
{
    [Theory]
    [InlineData("left-half")]
    [InlineData("right-half")]
    [InlineData("top-half")]
    [InlineData("bottom-half")]
    [InlineData("center")]
    [InlineData("full-screen")]
    public void Arrange_ValidPresets_Succeed(string preset)
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Act
        var result = _commands.Arrange(batch, preset);

        // Assert
        Assert.True(result.Success, $"Arrange '{preset}' failed: {result.ErrorMessage}");
        Assert.Equal("arrange", result.Action);
        Assert.Contains(preset, result.Message, StringComparison.OrdinalIgnoreCase);

        // Cleanup
        _commands.Hide(batch);
    }

    [Fact]
    public void Arrange_InvalidPreset_Throws()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Act & Assert
        Assert.ThrowsAny<Exception>(() => _commands.Arrange(batch, "invalid-preset"));
    }

    [Fact]
    public void Arrange_LeftHalf_PositionsCorrectly()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Act
        var result = _commands.Arrange(batch, "left-half");

        // Assert
        Assert.True(result.Success);

        var info = _commands.GetInfo(batch);
        Assert.True(info.IsVisible);
        Assert.Equal("normal", info.WindowState);
        // Excel COM positioning can have small offsets (e.g., -1.5 instead of 0)
        Assert.InRange(info.Left, -5, 5);
        Assert.InRange(info.Top, -5, 5);
        Assert.True(info.Width > 0);
        Assert.True(info.Height > 0);

        // Cleanup
        _commands.Hide(batch);
    }

    [Fact]
    public void Arrange_FullScreen_MaximizesWindow()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Act
        var result = _commands.Arrange(batch, "full-screen");

        // Assert
        Assert.True(result.Success);

        var info = _commands.GetInfo(batch);
        Assert.True(info.IsVisible);
        Assert.Equal("maximized", info.WindowState);

        // Cleanup
        _commands.Hide(batch);
    }

    [Fact]
    public void Arrange_WhenHidden_MakesVisible()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        _commands.Hide(batch);

        // Act
        var result = _commands.Arrange(batch, "center");

        // Assert
        Assert.True(result.Success);

        var info = _commands.GetInfo(batch);
        Assert.True(info.IsVisible, "Arrange should auto-show hidden Excel");

        // Cleanup
        _commands.Hide(batch);
    }
}
