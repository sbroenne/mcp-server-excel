// <copyright file="SheetCommandsTests.Outline.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Sheet;

/// <summary>
/// Integration tests for row and column grouping and worksheet outline controls.
/// </summary>
public partial class SheetCommandsTests
{
    [Fact]
    public void GroupAndUngroupRows_RoundTripsOutlineLevel()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        var group = _sheetCommands.Group(batch, sheetName, "2:5", OutlineAxis.Rows);
        var grouped = _sheetCommands.GetOutlineInfo(batch, sheetName, "2:5", OutlineAxis.Rows);

        Assert.True(group.Success, group.ErrorMessage);
        Assert.Equal(2, grouped.OutlineLevel);

        var ungroup = _sheetCommands.Ungroup(batch, sheetName, "2:5", OutlineAxis.Rows);
        var ungrouped = _sheetCommands.GetOutlineInfo(batch, sheetName, "2:5", OutlineAxis.Rows);

        Assert.True(ungroup.Success, ungroup.ErrorMessage);
        Assert.Equal(1, ungrouped.OutlineLevel);
    }

    [Fact]
    public void GroupAndUngroupColumns_RoundTripsOutlineLevel()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        var group = _sheetCommands.Group(batch, sheetName, "B:D", OutlineAxis.Columns);
        var grouped = _sheetCommands.GetOutlineInfo(batch, sheetName, "B:D", OutlineAxis.Columns);

        Assert.True(group.Success, group.ErrorMessage);
        Assert.Equal(2, grouped.OutlineLevel);

        var ungroup = _sheetCommands.Ungroup(batch, sheetName, "B:D", OutlineAxis.Columns);
        var ungrouped = _sheetCommands.GetOutlineInfo(batch, sheetName, "B:D", OutlineAxis.Columns);

        Assert.True(ungroup.Success, ungroup.ErrorMessage);
        Assert.Equal(1, ungrouped.OutlineLevel);
    }

    [Fact]
    public void SetOutlineSettings_RoundTripsSummaryPositionsAndStyles()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        var result = _sheetCommands.SetOutlineSettings(
            batch,
            sheetName,
            summaryRow: "above",
            summaryColumn: "left",
            automaticStyles: true);
        var info = _sheetCommands.GetOutlineInfo(batch, sheetName, "2:2", OutlineAxis.Rows);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal("above", info.SummaryRow);
        Assert.Equal("left", info.SummaryColumn);
        Assert.True(info.AutomaticStyles);
    }

    [Fact]
    public void SetOutlineSettings_InvalidColumn_DoesNotPartiallyChangeSummaryRow()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        _sheetCommands.SetOutlineSettings(
            batch,
            sheetName,
            summaryRow: "below",
            summaryColumn: "right",
            automaticStyles: false);

        Assert.Throws<ArgumentException>(() =>
            _sheetCommands.SetOutlineSettings(
                batch,
                sheetName,
                summaryRow: "above",
                summaryColumn: "invalid"));
        var info = _sheetCommands.GetOutlineInfo(batch, sheetName, "2:2", OutlineAxis.Rows);

        Assert.Equal("below", info.SummaryRow);
        Assert.Equal("right", info.SummaryColumn);
        Assert.False(info.AutomaticStyles);
    }

    [Fact]
    public void Group_InvalidAxis_DoesNotDefaultToRows()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            _sheetCommands.Group(batch, sheetName, "2:5", (OutlineAxis)0));
        var info = _sheetCommands.GetOutlineInfo(batch, sheetName, "2:5", OutlineAxis.Rows);

        Assert.Equal(1, info.OutlineLevel);
    }

    [Fact]
    public void ShowOutlineLevels_CollapsesGroupedRows()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        _sheetCommands.Group(batch, sheetName, "2:5", OutlineAxis.Rows);

        var result = _sheetCommands.ShowOutlineLevels(batch, sheetName, rowLevels: 1);
        var info = _sheetCommands.GetOutlineInfo(batch, sheetName, "2:5", OutlineAxis.Rows);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(info.Hidden);
    }

    [Fact]
    public void ClearOutline_RemovesRowAndColumnGroups()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        _sheetCommands.Group(batch, sheetName, "2:5", OutlineAxis.Rows);
        _sheetCommands.Group(batch, sheetName, "B:D", OutlineAxis.Columns);

        var result = _sheetCommands.ClearOutline(batch, sheetName);
        var rows = _sheetCommands.GetOutlineInfo(batch, sheetName, "2:5", OutlineAxis.Rows);
        var columns = _sheetCommands.GetOutlineInfo(batch, sheetName, "B:D", OutlineAxis.Columns);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(1, rows.OutlineLevel);
        Assert.Equal(1, columns.OutlineLevel);
    }
}
