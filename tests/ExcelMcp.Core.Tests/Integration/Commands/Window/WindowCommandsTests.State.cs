// <copyright file="WindowCommandsTests.State.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.Window;

/// <summary>
/// Tests for SetState and SetPosition operations.
/// </summary>
public partial class WindowCommandsTests
{
    [Theory]
    [InlineData("normal")]
    [InlineData("maximized")]
    [InlineData("minimized")]
    public void SetState_ValidStates_Succeed(string state)
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Act
        var result = _commands.SetState(batch, state);

        // Assert
        Assert.True(result.Success, $"SetState '{state}' failed: {result.ErrorMessage}");
        Assert.Equal("set-state", result.Action);
        Assert.Contains(state, result.Message, StringComparison.OrdinalIgnoreCase);

        // Cleanup
        _commands.Hide(batch);
    }

    [Fact]
    public void SetState_InvalidState_Throws()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Act & Assert
        Assert.ThrowsAny<Exception>(() => _commands.SetState(batch, "invalid-state"));
    }

    [Fact]
    public void SetPosition_AllParameters_UpdatesPosition()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Act
        var result = _commands.SetPosition(batch, left: 100, top: 50, width: 800, height: 600);

        // Assert
        Assert.True(result.Success, $"SetPosition failed: {result.ErrorMessage}");
        Assert.Equal("set-position", result.Action);

        // Verify position via GetInfo
        var info = _commands.GetInfo(batch);
        Assert.True(info.IsVisible, "SetPosition should make Excel visible");

        // Cleanup
        _commands.Hide(batch);
    }

    [Fact]
    public void SetPosition_PartialParameters_OnlyUpdatesProvided()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        _commands.Show(batch);
        _commands.SetState(batch, "normal");

        // Get initial position
        var beforeInfo = _commands.GetInfo(batch);

        // Act - only change left position
        var result = _commands.SetPosition(batch, left: 200);

        // Assert
        Assert.True(result.Success);

        var afterInfo = _commands.GetInfo(batch);
        Assert.Equal(200, afterInfo.Left, 1.0); // Allow small floating-point tolerance

        // Cleanup
        _commands.Hide(batch);
    }

    [Fact]
    public void SetState_MakesHiddenWindowVisible()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        _commands.Hide(batch);

        // Verify hidden
        var beforeInfo = _commands.GetInfo(batch);
        Assert.False(beforeInfo.IsVisible);

        // Act
        var result = _commands.SetState(batch, "normal");

        // Assert - should auto-show
        Assert.True(result.Success);
        var afterInfo = _commands.GetInfo(batch);
        Assert.True(afterInfo.IsVisible);

        // Cleanup
        _commands.Hide(batch);
    }
}
