using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceChartFormattingTests
{
    [Fact]
    public void List_EmptyWorkbook_ReturnsEmptyList()
    {
        // Act
        var batch = _fixture.BatchToken;
        var charts = _chartCommands.List(batch);

        // Assert
        RequireSuccess(charts);
        Assert.Empty(charts.Charts);
    }

    [Fact]
    public void CreateFromRange_ValidData_CreatesChart()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Act
        var createResult = _chartCommands.CreateFromRange(
            batch,
            _sheetName,
            "A1:B4",
            ChartType.ColumnClustered,
            100,
            50,
            400,
            300,
            "TestChart");

        // Assert
        RequireSuccess(createResult);
        Assert.Equal("TestChart", createResult.ChartName);
        Assert.Equal(_sheetName, createResult.SheetName);
        Assert.Equal(ChartType.ColumnClustered, createResult.ChartType);
        Assert.False(createResult.IsPivotChart);

        // Verify chart exists
        var charts = _chartCommands.List(batch);
        RequireSuccess(charts);
        var chart = Assert.Single(charts.Charts);
        Assert.Equal("TestChart", chart.Name);
        Assert.Equal(_sheetName, chart.SheetName);
        Assert.Equal(1, chart.SeriesCount);
        var read = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));
        Assert.Equal(100, read.Left);
        Assert.Equal(50, read.Top);
        Assert.Equal(400, read.Width);
        Assert.Equal(300, read.Height);
        AssertSeriesData(createResult.ChartName, 1, "Series1", ["A", "B", "C"], [10, 15, 20]);
    }

    [Fact]
    public void CreateFromTable_ValidTable_CreatesChart()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create test data and table
        RequireSuccess(_commands.SetValues(
            batch,
            _sheetName,
            "A1:B4",
            [["Category", "Values"],
             ["Q1", 100],
             ["Q2", 150],
             ["Q3", 200]], overwritePolicy: OverwritePolicy.Allow));
        RequireSuccess(_tableCommands.Create(
            batch,
            _sheetName,
            "SalesTable",
            "A1:B4",
            true));

        // Act
        var createResult = _chartCommands.CreateFromTable(
            batch,
            "SalesTable",
            _sheetName,
            ChartType.ColumnClustered,
            100,
            100,
            400,
            300,
            "TableChart");

        // Assert
        RequireSuccess(createResult);
        Assert.Equal("TableChart", createResult.ChartName);
        Assert.Equal(_sheetName, createResult.SheetName);
        Assert.Equal(ChartType.ColumnClustered, createResult.ChartType);
        Assert.False(createResult.IsPivotChart);

        // Verify chart exists
        var charts = _chartCommands.List(batch);
        RequireSuccess(charts);
        var chart = Assert.Single(charts.Charts);
        Assert.Equal("TableChart", chart.Name);
        AssertSeriesData(createResult.ChartName, 1, "Values", ["Q1", "Q2", "Q3"], [100, 150, 200]);
        RequireSuccess(_commands.SetValues(batch, _sheetName, "B3", [[175]], overwritePolicy: OverwritePolicy.Allow));
        AssertSeriesData(createResult.ChartName, 1, "Values", ["Q1", "Q2", "Q3"], [100, 175, 200]);
    }

    [Fact]
    public void CreateFromTable_NonExistentTable_ThrowsException()
    {
        // Act & Assert
        var batch = _fixture.BatchToken;
        var created = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Line));
        var before = RequireSuccess(_chartCommands.Read(batch, created.ChartName));
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.CreateFromTable(
                batch,
                "NonExistentTable",
                _sheetName,
                ChartType.ColumnClustered,
                50,
                50));

        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
        AssertChartUnchanged(before);
        var listed = RequireSuccess(_chartCommands.List(batch));
        Assert.Equal(created.ChartName, Assert.Single(listed.Charts).Name);
    }

    [Fact]
    public void CreateFromPivotTable_NonExistentPivotTable_ThrowsException()
    {
        // Act & Assert
        var batch = _fixture.BatchToken;
        var created = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Line));
        var before = RequireSuccess(_chartCommands.Read(batch, created.ChartName));
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.CreateFromPivotTable(
                batch,
                "NonExistentPivot",
                _sheetName,
                ChartType.ColumnClustered,
                50,
                50));

        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
        AssertChartUnchanged(before);
        var listed = RequireSuccess(_chartCommands.List(batch));
        Assert.Equal(created.ChartName, Assert.Single(listed.Charts).Name);
    }

    [Fact]
    public void Read_ExistingChart_ReturnsDetails()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Pie, 50, 50, 300, 300, "PieChart"));

        // Act
        var readResult = RequireSuccess(_chartCommands.Read(batch, "PieChart"));

        // Assert
        Assert.Equal("PieChart", readResult.Name);
        Assert.Equal(_sheetName, readResult.SheetName);
        Assert.Equal(ChartType.Pie, readResult.ChartType);
        Assert.False(readResult.IsPivotChart);
        Assert.Equal("Series1", Assert.Single(readResult.Series).Name);
        Assert.Equal(50, readResult.Left);
        Assert.Equal(50, readResult.Top);
        Assert.Equal(300, readResult.Width);
        Assert.Equal(300, readResult.Height);
        AssertSeriesData("PieChart", 1, "Series1", ["A", "B", "C"], [10, 15, 20]);
    }

    [Fact]
    public void Read_NonExistentChart_ReturnsError()
    {
        // Act & Assert
        var batch = _fixture.BatchToken;
        var created = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Line));
        var before = RequireSuccess(_chartCommands.Read(batch, created.ChartName));
        var exception = Assert.Throws<InvalidOperationException>(() => _chartCommands.Read(batch, "NonExistent"));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
        AssertChartUnchanged(before);
        AssertSeriesData(created.ChartName, 1, "Series1", ["A", "B", "C"], [10, 15, 20]);
    }

    [Fact]
    public void Delete_ExistingChart_RemovesChart()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B3", ChartType.Line, 50, 50));

        var chartsBefore = _chartCommands.List(batch);
        RequireSuccess(chartsBefore);
        var chartBefore = Assert.Single(chartsBefore.Charts);
        string chartName = chartBefore.Name;

        // Act
        RequireSuccess(_chartCommands.Delete(batch, chartName));

        // Assert - Verify chart removed
        var chartsAfter = _chartCommands.List(batch);
        RequireSuccess(chartsAfter);
        Assert.Empty(chartsAfter.Charts);
        var values = _commands.GetValues(batch, _sheetName, "A1:B3");
        RequireSuccess(values);
        Assert.Equal(["Category", "A", "B"], values.Values.Select(row => row[0]?.ToString()));
        Assert.Equal(["Series1", "10", "15"], values.Values.Select(row => row[1]?.ToString()));
    }

    [Fact]
    public void Move_ExistingChart_UpdatesPosition()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B3", ChartType.ColumnClustered, 100, 100, 300, 200));
        var untouched = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B3", ChartType.Line, 500, 100, 300, 200));
        var before = RequireSuccess(_chartCommands.Read(batch, untouched.ChartName));

        // Act
        RequireSuccess(_chartCommands.Move(batch, createResult.ChartName, left: 200, top: 150, width: 400, height: 250));

        // Assert - Verify position updated
        var readResult = RequireSuccess(_chartCommands.Read(batch, createResult.ChartName));
        Assert.Equal(200, readResult.Left);
        Assert.Equal(150, readResult.Top);
        Assert.Equal(400, readResult.Width);
        Assert.Equal(250, readResult.Height);
        AssertChartUnchanged(before);
        AssertSeriesData(createResult.ChartName, 1, "Series1", ["A", "B"], [10, 15]);
    }

    [Fact]
    public void List_MultipleCharts_ReturnsAll()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create multiple charts
        RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50, 300, 200, "Chart1"));
        RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Pie, 400, 50, 300, 200, "Chart2"));
        RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Line, 50, 300, 300, 200, "Chart3"));

        // Act
        var charts = _chartCommands.List(batch);

        // Assert
        RequireSuccess(charts);
        Assert.Equal(3, charts.Charts.Count);
        Assert.Contains(charts.Charts, c => c.Name == "Chart1" && c.ChartType == ChartType.ColumnClustered);
        Assert.Contains(charts.Charts, c => c.Name == "Chart2" && c.ChartType == ChartType.Pie);
        Assert.Contains(charts.Charts, c => c.Name == "Chart3" && c.ChartType == ChartType.Line);
        Assert.All(charts.Charts, chart =>
        {
            Assert.Equal(_sheetName, chart.SheetName);
            Assert.Equal(1, chart.SeriesCount);
            Assert.False(chart.IsPivotChart);
            AssertSeriesData(chart.Name, 1, "Series1", ["A", "B", "C"], [10, 15, 20]);
        });
    }

    [Fact]
    public void CreateFromRange_DifferentChartTypes_CreatesCorrectly()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Act & Assert - Test various chart types
        var chartTypes = new[]
        {
            ChartType.ColumnClustered,
            ChartType.BarClustered,
            ChartType.Line,
            ChartType.Pie,
            ChartType.XYScatter,
            ChartType.Area
        };

        int x = 50;
        foreach (var chartType in chartTypes)
        {
            var result = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:C5", chartType, x, 50, 250, 200));
            Assert.Equal(chartType, result.ChartType);
            Assert.Equal(chartType, InspectChart(result.ChartName, chart => (ChartType)chart.ChartType));
            object[] categories = chartType == ChartType.XYScatter ? [1, 2, 3, 4] : ["A", "B", "C", "D"];
            AssertSeriesData(result.ChartName, 1, "Series1", categories, [10, 15, 20, 25]);
            x += 300;
        }

        // Verify all created
        var charts = _chartCommands.List(batch);
        RequireSuccess(charts);
        Assert.Equal(chartTypes.Length, charts.Charts.Count);
    }
}
