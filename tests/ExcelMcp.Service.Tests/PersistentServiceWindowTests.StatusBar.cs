// <copyright file="WindowCommandsTests.StatusBar.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for SetStatusBar and ClearStatusBar operations.
/// </summary>
public sealed partial class PersistentServiceWindowTests
{
    [Fact]
    public void SetStatusBar_SetsText()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Act
        var result = _commands.SetStatusBar(batch, "Building PivotTable...");

        // Assert
        Assert.True(result.Success, $"SetStatusBar failed: {result.ErrorMessage}");
        Assert.Equal("set-status-bar", result.Action);
        Assert.Contains("Building PivotTable...", result.Message);

        // Cleanup
        _commands.ClearStatusBar(batch);
    }

    [Fact]
    public void ClearStatusBar_RestoresDefault()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        _commands.SetStatusBar(batch, "Some progress text");

        // Act
        var result = _commands.ClearStatusBar(batch);

        // Assert
        Assert.True(result.Success, $"ClearStatusBar failed: {result.ErrorMessage}");
        Assert.Equal("clear-status-bar", result.Action);
        Assert.Contains("default", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetStatusBar_Then_Clear_Roundtrip()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Act & Assert - Set
        var setResult = _commands.SetStatusBar(batch, "Step 1 of 3: Importing data...");
        Assert.True(setResult.Success);

        // Act & Assert - Update
        var updateResult = _commands.SetStatusBar(batch, "Step 2 of 3: Creating chart...");
        Assert.True(updateResult.Success);

        // Act & Assert - Clear
        var clearResult = _commands.ClearStatusBar(batch);
        Assert.True(clearResult.Success);
    }

    [Fact]
    public void SetStatusBar_MultipleUpdates_AllSucceed()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Act - Simulate progress updates
        for (int i = 1; i <= 5; i++)
        {
            var result = _commands.SetStatusBar(batch, $"Processing item {i} of 5...");
            Assert.True(result.Success, $"SetStatusBar call {i} failed: {result.ErrorMessage}");
        }

        // Cleanup
        _commands.ClearStatusBar(batch);
    }
}
