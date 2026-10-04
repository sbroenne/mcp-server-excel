// <copyright file="SheetCommandsTests.Outline.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for row and column grouping and worksheet outline controls.
/// </summary>
public sealed partial class PersistentServiceSheetTests
{
    [Fact]
    public void GroupAndUngroupRows_RoundTripsOutlineLevel()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var group = _sheetCommands.Group(batch, sheetName, "2:5", OutlineAxis.Rows);
        var grouped = _sheetCommands.GetOutlineInfo(batch, sheetName, "2:5", OutlineAxis.Rows);

        RequireSuccess(group);
        RequireSuccess(grouped);
        Assert.Equal(2, grouped.OutlineLevel);
        var untouched = _sheetCommands.GetOutlineInfo(batch, sheetName, "7:9", OutlineAxis.Rows);
        RequireSuccess(untouched);
        Assert.Equal(1, untouched.OutlineLevel);

        var ungroup = _sheetCommands.Ungroup(batch, sheetName, "2:5", OutlineAxis.Rows);
        var ungrouped = _sheetCommands.GetOutlineInfo(batch, sheetName, "2:5", OutlineAxis.Rows);

        RequireSuccess(ungroup);
        RequireSuccess(ungrouped);
        Assert.Equal(1, ungrouped.OutlineLevel);
    }

    [Fact]
    public void GroupAndUngroupColumns_RoundTripsOutlineLevel()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var group = _sheetCommands.Group(batch, sheetName, "B:D", OutlineAxis.Columns);
        var grouped = _sheetCommands.GetOutlineInfo(batch, sheetName, "B:D", OutlineAxis.Columns);

        RequireSuccess(group);
        RequireSuccess(grouped);
        Assert.Equal(2, grouped.OutlineLevel);
        var untouched = _sheetCommands.GetOutlineInfo(batch, sheetName, "F:H", OutlineAxis.Columns);
        RequireSuccess(untouched);
        Assert.Equal(1, untouched.OutlineLevel);

        var ungroup = _sheetCommands.Ungroup(batch, sheetName, "B:D", OutlineAxis.Columns);
        var ungrouped = _sheetCommands.GetOutlineInfo(batch, sheetName, "B:D", OutlineAxis.Columns);

        RequireSuccess(ungroup);
        RequireSuccess(ungrouped);
        Assert.Equal(1, ungrouped.OutlineLevel);
    }

    [Fact]
    public void SetOutlineSettings_RoundTripsSummaryPositionsAndStyles()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var result = _sheetCommands.SetOutlineSettings(
            batch,
            sheetName,
            summaryRow: "above",
            summaryColumn: "left",
            automaticStyles: true);
        var info = _sheetCommands.GetOutlineInfo(batch, sheetName, "2:2", OutlineAxis.Rows);

        RequireSuccess(result);
        RequireSuccess(info);
        Assert.Equal("above", info.SummaryRow);
        Assert.Equal("left", info.SummaryColumn);
        Assert.True(info.AutomaticStyles);
    }

    [Theory]
    [InlineData("above", "invalid", "Unknown summary column position")]
    [InlineData("invalid", "left", "Unknown summary row position")]
    public void SetOutlineSettings_InvalidPosition_DoesNotPartiallyChangeSettings(
        string row, string column, string expectedError)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var setup = _sheetCommands.SetOutlineSettings(
            batch,
            sheetName,
            summaryRow: "below",
            summaryColumn: "right",
            automaticStyles: false);
        RequireSuccess(setup);
        var before = _sheetCommands.GetOutlineInfo(batch, sheetName, "2:2", OutlineAxis.Rows);
        RequireSuccess(before);
        Assert.Equal("below", before.SummaryRow);
        Assert.Equal("right", before.SummaryColumn);
        Assert.False(before.AutomaticStyles);

        var exception = Assert.Throws<ArgumentException>(() =>
            _sheetCommands.SetOutlineSettings(
                batch,
                sheetName,
                summaryRow: row,
                summaryColumn: column,
                automaticStyles: true));
        var info = _sheetCommands.GetOutlineInfo(batch, sheetName, "2:2", OutlineAxis.Rows);

        Assert.Contains(expectedError, exception.Message, StringComparison.Ordinal);
        RequireSuccess(info);
        Assert.Equal("below", info.SummaryRow);
        Assert.Equal("right", info.SummaryColumn);
        Assert.False(info.AutomaticStyles);
    }

    [Fact]
    public void ShowOutlineLevels_CollapsesGroupedRows()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_sheetCommands.Group(batch, sheetName, "2:5", OutlineAxis.Rows));
        var before = _sheetCommands.GetOutlineInfo(batch, sheetName, "2:5", OutlineAxis.Rows);
        RequireSuccess(before);
        Assert.Equal(2, before.OutlineLevel);
        Assert.False(before.Hidden);

        var result = _sheetCommands.ShowOutlineLevels(batch, sheetName, rowLevels: 1);
        var info = _sheetCommands.GetOutlineInfo(batch, sheetName, "2:5", OutlineAxis.Rows);

        RequireSuccess(result);
        RequireSuccess(info);
        Assert.Equal(2, info.OutlineLevel);
        Assert.True(info.Hidden);
        var untouched = _sheetCommands.GetOutlineInfo(batch, sheetName, "7:9", OutlineAxis.Rows);
        RequireSuccess(untouched);
        Assert.False(untouched.Hidden);
        Assert.Equal(1, untouched.OutlineLevel);
    }

    [Fact]
    public void ClearOutline_RemovesRowAndColumnGroups()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_sheetCommands.Group(batch, sheetName, "2:5", OutlineAxis.Rows));
        RequireSuccess(_sheetCommands.Group(batch, sheetName, "B:D", OutlineAxis.Columns));
        var beforeRows = _sheetCommands.GetOutlineInfo(batch, sheetName, "2:5", OutlineAxis.Rows);
        var beforeColumns = _sheetCommands.GetOutlineInfo(batch, sheetName, "B:D", OutlineAxis.Columns);
        RequireSuccess(beforeRows);
        RequireSuccess(beforeColumns);
        Assert.Equal(2, beforeRows.OutlineLevel);
        Assert.Equal(2, beforeColumns.OutlineLevel);

        var result = _sheetCommands.ClearOutline(batch, sheetName);
        var rows = _sheetCommands.GetOutlineInfo(batch, sheetName, "2:5", OutlineAxis.Rows);
        var columns = _sheetCommands.GetOutlineInfo(batch, sheetName, "B:D", OutlineAxis.Columns);

        RequireSuccess(result);
        RequireSuccess(rows);
        RequireSuccess(columns);
        Assert.Equal(1, rows.OutlineLevel);
        Assert.Equal(1, columns.OutlineLevel);
    }

    [Fact]
    public void Group_InvalidAxis_DoesNotDefaultToRows()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_sheetCommands.Group(batch, sheetName, "2:5", OutlineAxis.Rows));
        RequireSuccess(_sheetCommands.Group(batch, sheetName, "B:D", OutlineAxis.Columns));
        var before = _sheetCommands.GetOutlineInfo(batch, sheetName, "2:5", OutlineAxis.Rows);
        RequireSuccess(before);
        Assert.Equal(2, before.OutlineLevel);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _sheetCommands.Group(
                batch,
                sheetName,
                "2:5",
                (OutlineAxis)0));
        Assert.Contains("axis", exception.Message, StringComparison.OrdinalIgnoreCase);
        var info = _sheetCommands.GetOutlineInfo(
            batch,
            sheetName,
            "2:5",
            OutlineAxis.Rows);

        RequireSuccess(info);
        Assert.Equal(2, info.OutlineLevel);
        Assert.Equal(before.Hidden, info.Hidden);
        var columns = _sheetCommands.GetOutlineInfo(batch, sheetName, "B:D", OutlineAxis.Columns);
        RequireSuccess(columns);
        Assert.Equal(2, columns.OutlineLevel);
    }
}
