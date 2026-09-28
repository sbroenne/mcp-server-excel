using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for Chart data source operations (SetSourceRange, AddSeries, RemoveSeries).
/// </summary>
public sealed partial class PersistentServiceChartFormattingTests
{
    [Fact]
    public void SetSourceRange_RegularChart_UpdatesDataSource()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Write additional data range for source change test
        _commands.SetValues(
            batch,
            _sheetName,
            "E1:F5",
            [["New", "Data"],
             ["X", 100],
             ["Y", 200],
             ["Z", 300],
             ["W", 400]]);

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50);

        // Act
        _chartCommands.SetSourceRange(batch, createResult.ChartName, "E1:F5");

        // Assert - Verify source range changed (Excel returns SERIES formula with Sheet1 reference)
        var readResult = _chartCommands.Read(batch, createResult.ChartName);
        Assert.Contains(_sheetName, readResult.SourceRange, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("$E$", readResult.SourceRange, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("$F$", readResult.SourceRange, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddSeries_RegularChart_AddsNewSeries()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Line, 50, 50);

        var readBefore = _chartCommands.Read(batch, createResult.ChartName);
        int initialSeriesCount = readBefore.Series.Count;

        // Act
        var addSeriesResult = _chartCommands.AddSeries(
            batch,
            createResult.ChartName,
            "NewSeries",
            $"{_sheetName}!C2:C4",
            $"{_sheetName}!A2:A4");

        // Assert
        Assert.Equal("NewSeries", addSeriesResult.Name);

        // Verify series added
        var readAfter = _chartCommands.Read(batch, createResult.ChartName);
        Assert.Equal(initialSeriesCount + 1, readAfter.Series.Count);
        Assert.Contains(readAfter.Series, s => s.Name == "NewSeries");
    }

    [Fact]
    public void RemoveSeries_RegularChart_RemovesSeries()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:C4", ChartType.ColumnClustered, 50, 50);

        var readBefore = _chartCommands.Read(batch, createResult.ChartName);
        int initialSeriesCount = readBefore.Series.Count;
        Assert.True(initialSeriesCount >= 2, "Need at least 2 series for test");

        // Act - Remove first series (index 1)
        _chartCommands.RemoveSeries(batch, createResult.ChartName, 1);

        // Assert - Verify series removed
        var readAfter = _chartCommands.Read(batch, createResult.ChartName);
        Assert.Equal(initialSeriesCount - 1, readAfter.Series.Count);
    }

    [Fact]
    public void RemoveSeries_InvalidSeriesIndex_ThrowsHelpfulPublicError()
    {
        var batch = _fixture.BatchToken;
        var createResult = _chartCommands.CreateFromRange(
            batch,
            _sheetName,
            "A1:B4",
            ChartType.Line,
            50,
            50);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.RemoveSeries(
                batch,
                createResult.ChartName,
                999));

        Assert.Contains(
            "series",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddSeries_WithoutCategoryRange_CreatesSeriesSuccessfully()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.XYScatter, 50, 50);

        // Act - Add series without category range
        var addResult = _chartCommands.AddSeries(batch, createResult.ChartName, "Series3", $"{_sheetName}!C2:C4", null);

        // Assert
        Assert.Equal("Series3", addResult.Name);
    }

    [Fact]
    public void SetSourceRange_NonExistentChart_ReturnsError()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Act & Assert
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.SetSourceRange(batch, "NonExistent", "A1:B10"));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetSourceRange_ExpandedRange_UpdatesChartData()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B3", ChartType.BarClustered, 50, 50);

        // Act - Expand to include more rows
        _chartCommands.SetSourceRange(batch, createResult.ChartName, "A1:B5");

        // Assert - Verify expanded range (Excel returns SERIES formula with Sheet1 reference)
        var readResult = _chartCommands.Read(batch, createResult.ChartName);
        Assert.Contains(_sheetName, readResult.SourceRange, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("$A$", readResult.SourceRange, StringComparison.OrdinalIgnoreCase);
        // Verify it includes row 5 (expanded from 3 to 5)
        Assert.Contains("$5", readResult.SourceRange, StringComparison.OrdinalIgnoreCase);
    }
}


