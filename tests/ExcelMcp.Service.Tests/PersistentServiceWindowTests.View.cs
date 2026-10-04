// <copyright file="WindowCommandsTests.View.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for worksheet-specific window view operations.
/// </summary>
public sealed partial class PersistentServiceWindowTests
{
    [Fact]
    public void FreezeAndUnfreezePanes_RoundTripsViewState()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var freeze = _commands.FreezePanes(batch, sheetName, frozenRows: 2, frozenColumns: 1);
        var frozenView = _commands.GetView(batch, sheetName);

        Assert.True(freeze.Success, freeze.ErrorMessage);
        Assert.True(frozenView.Success, frozenView.ErrorMessage);
        Assert.True(frozenView.FreezePanes);
        Assert.Equal(2, frozenView.SplitRow);
        Assert.Equal(1, frozenView.SplitColumn);

        var unfreeze = _commands.UnfreezePanes(batch, sheetName);
        var unfrozenView = _commands.GetView(batch, sheetName);

        Assert.True(unfreeze.Success, unfreeze.ErrorMessage);
        Assert.True(unfrozenView.Success, unfrozenView.ErrorMessage);
        Assert.False(unfrozenView.FreezePanes);
        Assert.Equal(0, unfrozenView.SplitRow);
        Assert.Equal(0, unfrozenView.SplitColumn);
    }

    [Fact]
    public void SetSplit_ReplacesFrozenPanesWithMovableSplit()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.FreezePanes(batch, sheetName, frozenRows: 2, frozenColumns: 1).Success);
        var frozen = _commands.GetView(batch, sheetName);
        Assert.True(frozen.Success, frozen.ErrorMessage);
        Assert.True(frozen.FreezePanes);
        Assert.Equal(2, frozen.SplitRow);
        Assert.Equal(1, frozen.SplitColumn);

        var split = _commands.SetSplit(batch, sheetName, splitRows: 4, splitColumns: 2);
        var view = _commands.GetView(batch, sheetName);

        Assert.True(split.Success, split.ErrorMessage);
        Assert.True(view.Success, view.ErrorMessage);
        Assert.False(view.FreezePanes);
        Assert.Equal(4, view.SplitRow);
        Assert.Equal(2, view.SplitColumn);
    }

    [Fact]
    public void SetZoom_UpdatesTargetWorksheetView()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var otherSheet = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetZoom(batch, sheetName, 70).Success);
        Assert.True(_commands.SetZoom(batch, otherSheet, 85).Success);
        var before = _commands.GetView(batch, sheetName);
        Assert.True(before.Success, before.ErrorMessage);
        Assert.Equal(70, before.Zoom);

        var result = _commands.SetZoom(batch, sheetName, 135);
        var view = _commands.GetView(batch, sheetName);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(view.Success, view.ErrorMessage);
        Assert.Equal(135, view.Zoom);
        var untouched = _commands.GetView(batch, otherSheet);
        Assert.True(untouched.Success, untouched.ErrorMessage);
        Assert.Equal(85, untouched.Zoom);
    }

    [Fact]
    public void SetDisplayOptions_UpdatesGridlinesHeadingsAndOutlineSymbols()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetDisplayOptions(batch, sheetName,
            showGridlines: true, showHeadings: true, showOutlineSymbols: true).Success);
        var before = _commands.GetView(batch, sheetName);
        Assert.True(before.Success, before.ErrorMessage);
        Assert.True(before.DisplayGridlines);
        Assert.True(before.DisplayHeadings);
        Assert.True(before.DisplayOutlineSymbols);

        var result = _commands.SetDisplayOptions(
            batch,
            sheetName,
            showGridlines: false,
            showHeadings: false,
            showOutlineSymbols: false);
        var view = _commands.GetView(batch, sheetName);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(view.Success, view.ErrorMessage);
        Assert.False(view.DisplayGridlines);
        Assert.False(view.DisplayHeadings);
        Assert.False(view.DisplayOutlineSymbols);
    }

    [Fact]
    public void SetDisplayOptions_ShowFormulasRoundTripsAndLeavesOtherOptionsUnchanged()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var initialView = _commands.GetView(batch, sheetName);
        Assert.True(initialView.Success, initialView.ErrorMessage);

        var show = _commands.SetDisplayOptions(batch, sheetName, showFormulas: true);
        var formulasShown = _commands.GetView(batch, sheetName);

        Assert.True(show.Success, show.ErrorMessage);
        Assert.True(formulasShown.Success, formulasShown.ErrorMessage);
        Assert.True(formulasShown.DisplayFormulas);
        Assert.Equal(initialView.DisplayGridlines, formulasShown.DisplayGridlines);
        Assert.Equal(initialView.DisplayHeadings, formulasShown.DisplayHeadings);
        Assert.Equal(initialView.DisplayOutlineSymbols, formulasShown.DisplayOutlineSymbols);

        var hide = _commands.SetDisplayOptions(batch, sheetName, showFormulas: false);
        var formulasHidden = _commands.GetView(batch, sheetName);

        Assert.True(hide.Success, hide.ErrorMessage);
        Assert.True(formulasHidden.Success, formulasHidden.ErrorMessage);
        Assert.False(formulasHidden.DisplayFormulas);
        Assert.Equal(initialView.DisplayGridlines, formulasHidden.DisplayGridlines);
        Assert.Equal(initialView.DisplayHeadings, formulasHidden.DisplayHeadings);
        Assert.Equal(initialView.DisplayOutlineSymbols, formulasHidden.DisplayOutlineSymbols);
    }

    [Fact]
    public void FreezePanes_WithoutRowsOrColumns_ReturnsFailure()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.FreezePanes(batch, sheetName, frozenRows: 2, frozenColumns: 1).Success);
        var before = _commands.GetView(batch, sheetName);
        Assert.True(before.Success, before.ErrorMessage);
        Assert.True(before.FreezePanes);
        Assert.Equal(2, before.SplitRow);
        Assert.Equal(1, before.SplitColumn);

        var exception = Assert.Throws<ArgumentException>(() => _commands.FreezePanes(batch, sheetName));

        Assert.Contains("row", exception.Message, StringComparison.OrdinalIgnoreCase);
        var after = _commands.GetView(batch, sheetName);
        Assert.True(after.Success, after.ErrorMessage);
        Assert.True(after.FreezePanes);
        Assert.Equal(2, after.SplitRow);
        Assert.Equal(1, after.SplitColumn);
    }

    [Theory]
    [InlineData(-1, 1)]
    [InlineData(1, -1)]
    [InlineData(1_048_576, 1)]
    [InlineData(1, 16_384)]
    public void FreezePanes_InvalidCounts_PreservesExistingPaneState(int rows, int columns)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.FreezePanes(batch, sheetName, frozenRows: 1, frozenColumns: 1).Success);
        var before = _commands.GetView(batch, sheetName);
        Assert.True(before.Success, before.ErrorMessage);
        Assert.True(before.FreezePanes);
        Assert.Equal(1, before.SplitRow);
        Assert.Equal(1, before.SplitColumn);

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            _commands.FreezePanes(batch, sheetName, frozenRows: rows, frozenColumns: columns));
        var view = _commands.GetView(batch, sheetName);

        Assert.True(view.Success, view.ErrorMessage);
        Assert.True(view.FreezePanes);
        Assert.Equal(1, view.SplitRow);
        Assert.Equal(1, view.SplitColumn);
    }

    [Theory]
    [InlineData(-1, 1)]
    [InlineData(1, -1)]
    [InlineData(1_048_576, 1)]
    [InlineData(1, 16_384)]
    public void SetSplit_InvalidCounts_PreservesExistingPaneState(int rows, int columns)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetSplit(batch, sheetName, splitRows: 2, splitColumns: 2).Success);
        var before = _commands.GetView(batch, sheetName);
        Assert.True(before.Success, before.ErrorMessage);
        Assert.False(before.FreezePanes);
        Assert.Equal(2, before.SplitRow);
        Assert.Equal(2, before.SplitColumn);

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            _commands.SetSplit(batch, sheetName, splitRows: rows, splitColumns: columns));
        var view = _commands.GetView(batch, sheetName);

        Assert.True(view.Success, view.ErrorMessage);
        Assert.False(view.FreezePanes);
        Assert.Equal(2, view.SplitRow);
        Assert.Equal(2, view.SplitColumn);
    }

    [Theory]
    [InlineData(9)]
    [InlineData(401)]
    public void SetZoom_InvalidValue_PreservesExistingView(int zoom)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetZoom(batch, sheetName, 135).Success);
        Assert.True(_commands.FreezePanes(batch, sheetName, frozenRows: 2, frozenColumns: 1).Success);
        var before = _commands.GetView(batch, sheetName);
        Assert.True(before.Success, before.ErrorMessage);
        Assert.Equal(135, before.Zoom);

        var exception = Assert.Throws<ArgumentOutOfRangeException>(() => _commands.SetZoom(batch, sheetName, zoom));

        Assert.Contains("Zoom must be between 10 and 400 percent", exception.Message, StringComparison.Ordinal);
        var after = _commands.GetView(batch, sheetName);
        Assert.True(after.Success, after.ErrorMessage);
        Assert.Equal(before.Zoom, after.Zoom);
        Assert.Equal(before.FreezePanes, after.FreezePanes);
        Assert.Equal(before.SplitRow, after.SplitRow);
        Assert.Equal(before.SplitColumn, after.SplitColumn);
        Assert.Equal(before.DisplayGridlines, after.DisplayGridlines);
        Assert.Equal(before.DisplayHeadings, after.DisplayHeadings);
        Assert.Equal(before.DisplayOutlineSymbols, after.DisplayOutlineSymbols);
        Assert.Equal(before.DisplayFormulas, after.DisplayFormulas);
    }
}
