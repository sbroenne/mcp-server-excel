using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for Chart data source operations (SetSourceRange, AddSeries, RemoveSeries).
/// </summary>
public sealed partial class PersistentServiceChartFormattingTests
{
    private static readonly string[] ReadbackQuarters = ["Q1", "Q2"];
    private static readonly string[] ReadbackSeriesNames = ["Alpha", "Beta", "Gamma"];

    [Fact]
    public void ReadAndList_RegularChart_ReturnActualValuesAndCategoriesForEverySeries()
    {
        var batch = _fixture.BatchToken;
        RequireSuccess(_commands.SetValues(batch, _sheetName, "E1:H3",
        [
            ["Quarter", "Alpha", "Beta", "Gamma"],
            ["Q1", 10, 20, 30],
            ["Q2", 40, 50, 60]
        ]));
        var created = _chartCommands.CreateFromRange(batch, _sheetName, "E1:H3", ChartType.Line);
        RequireSuccess(created);
        var read = _chartCommands.Read(batch, created.ChartName);
        RequireSuccess(read);
        Assert.False(read.IsPivotChart);
        Assert.Equal(ReadbackSeriesNames, read.Series.Select(series => series.Name));
        var expected = new[] { new[] { 10d, 40d }, new[] { 20d, 50d }, new[] { 30d, 60d } };
        for (var index = 0; index < expected.Length; index++)
        {
            Assert.Equal(expected[index], read.Series[index].Values.Select(value =>
                Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture)));
            Assert.Equal(ReadbackQuarters, read.Series[index].Categories.Select(value => value?.ToString()));
            Assert.Equal(string.Empty, read.Series[index].ValuesRange);
            Assert.Null(read.Series[index].CategoryRange);
        }
        var list = _chartCommands.List(batch);
        RequireSuccess(list);
        Assert.Equal(3, Assert.Single(list.Charts, chart => chart.Name == created.ChartName).SeriesCount);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task AddSeries_MissingSource_PreservesExistingSeries(bool missingValues)
    {
        var batch = _fixture.BatchToken;
        var created = RequireSuccess(_chartCommands.CreateFromRange(
            batch, _sheetName, "A1:B4", ChartType.Line));
        var before = RequireSuccess(_chartCommands.Read(batch, created.ChartName));
        var rejected = await _fixture.SendForFailureAsync("chartconfig.add-series", new
        {
            chartName = created.ChartName,
            seriesName = "RejectedSeries",
            valuesRange = missingValues ? "MissingSource!$C$2:$C$4" : $"{_sheetName}!$C$2:$C$4",
            categoryRange = missingValues ? $"{_sheetName}!$A$2:$A$4" : "MissingSource!$A$2:$A$4"
        });

        Assert.Equal("ComInterop", rejected.ErrorCategory);
        Assert.Contains("COMException", rejected.ErrorMessage);
        AssertChartUnchanged(before);
        AssertSeriesData(created.ChartName, 1, "Series1", ["A", "B", "C"], [10, 15, 20]);
    }

    [Fact]
    public void SetSourceRange_RegularChart_UpdatesDataSource()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Write additional data range for source change test
        RequireSuccess(_commands.SetValues(
            batch,
            _sheetName,
            "E1:F5",
            [["New", "Data"],
             ["X", 100],
             ["Y", 200],
             ["Z", 300],
             ["W", 400]]));

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50));
        AssertSeriesData(createResult.ChartName, 1, "Series1", ["A", "B", "C"], [10, 15, 20]);

        // Act
        RequireSuccess(_chartCommands.SetSourceRange(batch, createResult.ChartName, "E1:F5"));

        // Assert - Verify source range changed (Excel returns SERIES formula with Sheet1 reference)
        var readResult = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));
        Assert.Contains(_sheetName, readResult.SourceRange, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("$E$", readResult.SourceRange, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("$F$", readResult.SourceRange, StringComparison.OrdinalIgnoreCase);
        Assert.Single(readResult.Series);
        AssertSeriesData(createResult.ChartName, 1, "Data", ["X", "Y", "Z", "W"], [100, 200, 300, 400]);
    }

    [Fact]
    public void AddSeries_RegularChart_AddsNewSeries()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Line, 50, 50));

        var readBefore = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));
        int initialSeriesCount = readBefore.Series.Count;
        Assert.Equal(1, initialSeriesCount);
        AssertSeriesData(createResult.ChartName, 1, "Series1", ["A", "B", "C"], [10, 15, 20]);

        // Act
        var addSeriesResult = _chartCommands.AddSeries(
            batch,
            createResult.ChartName,
            "NewSeries",
            $"{_sheetName}!C2:C4",
            $"{_sheetName}!A2:A4");

        // Assert
        Assert.Equal("NewSeries", addSeriesResult.Name);
        Assert.Equal($"{_sheetName}!C2:C4", addSeriesResult.ValuesRange);
        Assert.Equal($"{_sheetName}!A2:A4", addSeriesResult.CategoryRange);

        // Verify series added
        var readAfter = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));
        Assert.Equal(initialSeriesCount + 1, readAfter.Series.Count);
        Assert.Contains(readAfter.Series, s => s.Name == "NewSeries");
        AssertSeriesData(createResult.ChartName, 1, "Series1", ["A", "B", "C"], [10, 15, 20]);
        AssertSeriesData(createResult.ChartName, 2, "NewSeries", ["A", "B", "C"], [20, 25, 30]);
        RequireSuccess(_commands.SetValues(batch, _sheetName, "C3", [[26]], overwritePolicy: OverwritePolicy.Allow));
        AssertSeriesData(createResult.ChartName, 2, "NewSeries", ["A", "B", "C"], [20, 26, 30]);
        AssertSeriesData(createResult.ChartName, 1, "Series1", ["A", "B", "C"], [10, 15, 20]);
    }

    [Fact]
    public void RemoveSeries_RegularChart_RemovesSeries()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:C4", ChartType.ColumnClustered, 50, 50));

        var readBefore = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));
        int initialSeriesCount = readBefore.Series.Count;
        Assert.Equal(2, initialSeriesCount);
        AssertSeriesData(createResult.ChartName, 1, "Series1", ["A", "B", "C"], [10, 15, 20]);
        AssertSeriesData(createResult.ChartName, 2, "Series2", ["A", "B", "C"], [20, 25, 30]);

        // Act - Remove first series (index 1)
        RequireSuccess(_chartCommands.RemoveSeries(batch, createResult.ChartName, 1));

        // Assert - Verify series removed
        var readAfter = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));
        Assert.Equal(initialSeriesCount - 1, readAfter.Series.Count);
        Assert.Equal("Series2", Assert.Single(readAfter.Series).Name);
        AssertSeriesData(createResult.ChartName, 1, "Series2", ["A", "B", "C"], [20, 25, 30]);
    }

    [Fact]
    public void RemoveSeries_InvalidSeriesIndex_ThrowsHelpfulPublicError()
    {
        var batch = _fixture.BatchToken;
        var createResult = RequireSuccess(_chartCommands.CreateFromRange(
            batch,
            _sheetName,
            "A1:B4",
            ChartType.Line,
            50,
            50));
        var before = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.RemoveSeries(
                batch,
                createResult.ChartName,
                999));

        Assert.Contains(
            "series",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
        AssertChartUnchanged(before);
        AssertSeriesData(createResult.ChartName, 1, "Series1", ["A", "B", "C"], [10, 15, 20]);
    }

    [Fact]
    public void AddSeries_WithoutCategoryRange_CreatesSeriesSuccessfully()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.XYScatter, 50, 50));

        // Act - Add series without category range
        var addResult = _chartCommands.AddSeries(batch, createResult.ChartName, "Series3", $"{_sheetName}!C2:C4", null);

        // Assert
        Assert.Equal("Series3", addResult.Name);
        Assert.Equal($"{_sheetName}!C2:C4", addResult.ValuesRange);
        Assert.Null(addResult.CategoryRange);
        var read = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));
        Assert.Equal(["Series1", "Series3"], read.Series.Select(series => series.Name));
        AssertSeriesData(createResult.ChartName, 2, "Series3", [1, 2, 3], [20, 25, 30]);
    }

    [Fact]
    public void SetSourceRange_NonExistentChart_ReturnsError()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var created = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Line));
        var before = RequireSuccess(_chartCommands.Read(batch, created.ChartName));

        // Act & Assert
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.SetSourceRange(batch, "NonExistent", "A1:B10"));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
        AssertChartUnchanged(before);
        AssertSeriesData(created.ChartName, 1, "Series1", ["A", "B", "C"], [10, 15, 20]);
    }

    [Fact]
    public void SetSourceRange_ExpandedRange_UpdatesChartData()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B3", ChartType.BarClustered, 50, 50));
        AssertSeriesData(createResult.ChartName, 1, "Series1", ["A", "B"], [10, 15]);

        // Act - Expand to include more rows
        RequireSuccess(_chartCommands.SetSourceRange(batch, createResult.ChartName, "A1:B5"));

        // Assert - Verify expanded range (Excel returns SERIES formula with Sheet1 reference)
        var readResult = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));
        Assert.Contains(_sheetName, readResult.SourceRange, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("$A$", readResult.SourceRange, StringComparison.OrdinalIgnoreCase);
        // Verify it includes row 5 (expanded from 3 to 5)
        Assert.Contains("$5", readResult.SourceRange, StringComparison.OrdinalIgnoreCase);
        Assert.Single(readResult.Series);
        AssertSeriesData(createResult.ChartName, 1, "Series1", ["A", "B", "C", "D"], [10, 15, 20, 25]);
    }
}
