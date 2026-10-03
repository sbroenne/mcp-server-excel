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
    [Theory]
    [InlineData("null")]
    [InlineData("missing")]
    public void ClearStatusBar_NativeResetControl_ReturnsBooleanFalse(string reset)
    {
        _fixture.ExecuteRawVerification((context, _) =>
        {
            context.App.StatusBar = "Native reset control";
            Assert.Equal("Native reset control", context.App.StatusBar);
            context.App.StatusBar = reset switch
            {
                "null" => null!,
                "missing" => Type.Missing,
                _ => throw new ArgumentOutOfRangeException(nameof(reset))
            };
            Assert.False(Assert.IsType<bool>(context.App.StatusBar));
        });
    }

    [Fact]
    public void SetStatusBar_Then_Clear_Roundtrip()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Act & Assert - Set
        var setResult = _commands.SetStatusBar(batch, "Step 1 of 3: Importing data...");
        RequireSuccess(setResult);
        Assert.Equal("set-status-bar", setResult.Action);
        Assert.Contains("Step 1 of 3: Importing data...", setResult.Message);
        Assert.Equal("Step 1 of 3: Importing data...", ReadStatusBar());

        // Act & Assert - Update
        var updateResult = _commands.SetStatusBar(batch, "Step 2 of 3: Creating chart...");
        RequireSuccess(updateResult);
        Assert.Equal("set-status-bar", updateResult.Action);
        Assert.Contains("Step 2 of 3: Creating chart...", updateResult.Message);
        Assert.Equal("Step 2 of 3: Creating chart...", ReadStatusBar());

        // Act & Assert - Clear
        var clearResult = _commands.ClearStatusBar(batch);
        RequireSuccess(clearResult);
        Assert.Equal("clear-status-bar", clearResult.Action);
        Assert.Contains("default", clearResult.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(Assert.IsType<bool>(ReadStatusBar()));
    }

    private object ReadStatusBar() =>
        _fixture.ExecuteRawVerification((ctx, ct) => ctx.App.StatusBar);
}
