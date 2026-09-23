// <copyright file="WindowCommandsTests.View.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.Window;

/// <summary>
/// Integration tests for worksheet-specific window view operations.
/// </summary>
public partial class WindowCommandsTests
{
    [Fact]
    public void FreezeAndUnfreezePanes_RoundTripsViewState()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        const string sheetName = "ViewFreeze";
        new SheetCommands().Create(batch, sheetName);

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
        Assert.False(unfrozenView.FreezePanes);
        Assert.Equal(0, unfrozenView.SplitRow);
        Assert.Equal(0, unfrozenView.SplitColumn);
    }

    [Fact]
    public void SetSplit_ReplacesFrozenPanesWithMovableSplit()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        const string sheetName = "ViewSplit";
        new SheetCommands().Create(batch, sheetName);
        _commands.FreezePanes(batch, sheetName, frozenRows: 2, frozenColumns: 1);
        _commands.GetView(batch, sheetName);

        var split = _commands.SetSplit(batch, sheetName, splitRows: 4, splitColumns: 2);
        var view = _commands.GetView(batch, sheetName);

        Assert.True(split.Success, split.ErrorMessage);
        Assert.False(view.FreezePanes);
        Assert.Equal(4, view.SplitRow);
        Assert.Equal(2, view.SplitColumn);
    }

    [Fact]
    public void SetZoom_UpdatesTargetWorksheetView()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        const string sheetName = "ViewZoom";
        new SheetCommands().Create(batch, sheetName);

        var result = _commands.SetZoom(batch, sheetName, 135);
        var view = _commands.GetView(batch, sheetName);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(135, view.Zoom);
    }

    [Fact]
    public void SetDisplayOptions_UpdatesGridlinesHeadingsAndOutlineSymbols()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        const string sheetName = "ViewDisplay";
        new SheetCommands().Create(batch, sheetName);

        var result = _commands.SetDisplayOptions(
            batch,
            sheetName,
            showGridlines: false,
            showHeadings: false,
            showOutlineSymbols: false);
        var view = _commands.GetView(batch, sheetName);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.False(view.DisplayGridlines);
        Assert.False(view.DisplayHeadings);
        Assert.False(view.DisplayOutlineSymbols);
    }

    [Fact]
    public void SetDisplayOptions_ShowFormulasRoundTripsAndLeavesOtherOptionsUnchanged()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        const string sheetName = "ViewFormulas";
        new SheetCommands().Create(batch, sheetName);
        var initialView = _commands.GetView(batch, sheetName);

        var show = _commands.SetDisplayOptions(batch, sheetName, showFormulas: true);
        var formulasShown = _commands.GetView(batch, sheetName);

        Assert.True(show.Success, show.ErrorMessage);
        Assert.True(formulasShown.DisplayFormulas);
        Assert.Equal(initialView.DisplayGridlines, formulasShown.DisplayGridlines);
        Assert.Equal(initialView.DisplayHeadings, formulasShown.DisplayHeadings);
        Assert.Equal(initialView.DisplayOutlineSymbols, formulasShown.DisplayOutlineSymbols);

        var hide = _commands.SetDisplayOptions(batch, sheetName, showFormulas: false);
        var formulasHidden = _commands.GetView(batch, sheetName);

        Assert.True(hide.Success, hide.ErrorMessage);
        Assert.False(formulasHidden.DisplayFormulas);
    }

    [Fact]
    public void FreezePanes_WithoutRowsOrColumns_ReturnsFailure()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        const string sheetName = "ViewInvalid";
        new SheetCommands().Create(batch, sheetName);

        var exception = Assert.Throws<ArgumentException>(() => _commands.FreezePanes(batch, sheetName));

        Assert.Contains("row", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void FreezePanes_AboveWorksheetLimits_PreservesExistingPaneState()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        const string sheetName = "ViewFreezeLimit";
        new SheetCommands().Create(batch, sheetName);
        _commands.FreezePanes(batch, sheetName, frozenRows: 1, frozenColumns: 1);

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            _commands.FreezePanes(batch, sheetName, frozenRows: 1_048_576));
        var view = _commands.GetView(batch, sheetName);

        Assert.True(view.FreezePanes);
        Assert.Equal(1, view.SplitRow);
        Assert.Equal(1, view.SplitColumn);
    }

    [Fact]
    public void SetSplit_AboveWorksheetLimits_PreservesExistingPaneState()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        const string sheetName = "ViewSplitLimit";
        new SheetCommands().Create(batch, sheetName);
        _commands.SetSplit(batch, sheetName, splitRows: 2, splitColumns: 2);

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            _commands.SetSplit(batch, sheetName, splitColumns: 16_384));
        var view = _commands.GetView(batch, sheetName);

        Assert.False(view.FreezePanes);
        Assert.Equal(2, view.SplitRow);
        Assert.Equal(2, view.SplitColumn);
    }
}
