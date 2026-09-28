using Sbroenne.ExcelMcp.Core.Commands.Chart;
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
        Assert.True(charts.Success);
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
        Assert.Equal("TestChart", createResult.ChartName);
        Assert.Equal(_sheetName, createResult.SheetName);
        Assert.Equal(ChartType.ColumnClustered, createResult.ChartType);
        Assert.False(createResult.IsPivotChart);

        // Verify chart exists
        var charts = _chartCommands.List(batch);
        Assert.True(charts.Success);
        var chart = Assert.Single(charts.Charts);
        Assert.Equal("TestChart", chart.Name);
    }

    [Fact]
    public void CreateFromTable_ValidTable_CreatesChart()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create test data and table
        _commands.SetValues(
            batch,
            _sheetName,
            "A1:B4",
            [["Category", "Values"],
             ["Q1", 100],
             ["Q2", 150],
             ["Q3", 200]]);
        _tableCommands.Create(
            batch,
            _sheetName,
            "SalesTable",
            "A1:B4",
            true);

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
        Assert.True(createResult.Success, $"CreateFromTable failed: {createResult.ChartName}");
        Assert.Equal("TableChart", createResult.ChartName);
        Assert.Equal(_sheetName, createResult.SheetName);
        Assert.Equal(ChartType.ColumnClustered, createResult.ChartType);
        Assert.False(createResult.IsPivotChart);

        // Verify chart exists
        var charts = _chartCommands.List(batch);
        Assert.True(charts.Success);
        var chart = Assert.Single(charts.Charts);
        Assert.Equal("TableChart", chart.Name);
    }

    [Fact]
    public void CreateFromTable_NonExistentTable_ThrowsException()
    {
        // Act & Assert
        var batch = _fixture.BatchToken;
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.CreateFromTable(
                batch,
                "NonExistentTable",
                _sheetName,
                ChartType.ColumnClustered,
                50,
                50));

        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void CreateFromPivotTable_NonExistentPivotTable_ThrowsException()
    {
        // Act & Assert
        var batch = _fixture.BatchToken;
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.CreateFromPivotTable(
                batch,
                "NonExistentPivot",
                _sheetName,
                ChartType.ColumnClustered,
                50,
                50));

        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Read_ExistingChart_ReturnsDetails()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Pie, 50, 50, 300, 300, "PieChart");

        // Act
        var readResult = _chartCommands.Read(batch, "PieChart");

        // Assert
        Assert.Equal("PieChart", readResult.Name);
        Assert.Equal(_sheetName, readResult.SheetName);
        Assert.Equal(ChartType.Pie, readResult.ChartType);
        Assert.False(readResult.IsPivotChart);
        Assert.True(readResult.Series.Count > 0);
    }

    [Fact]
    public void Read_NonExistentChart_ReturnsError()
    {
        // Act & Assert
        var batch = _fixture.BatchToken;
        var exception = Assert.Throws<InvalidOperationException>(() => _chartCommands.Read(batch, "NonExistent"));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Delete_ExistingChart_RemovesChart()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        _chartCommands.CreateFromRange(batch, _sheetName, "A1:B3", ChartType.Line, 50, 50);

        var chartsBefore = _chartCommands.List(batch);
        Assert.True(chartsBefore.Success);
        var chartBefore = Assert.Single(chartsBefore.Charts);
        string chartName = chartBefore.Name;

        // Act
        _chartCommands.Delete(batch, chartName);

        // Assert - Verify chart removed
        var chartsAfter = _chartCommands.List(batch);
        Assert.True(chartsAfter.Success);
        Assert.Empty(chartsAfter.Charts);
    }

    [Fact]
    public void Move_ExistingChart_UpdatesPosition()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _chartCommands.CreateFromRange(batch, _sheetName, "A1:B3", ChartType.ColumnClustered, 100, 100, 300, 200);

        // Act
        _chartCommands.Move(batch, createResult.ChartName, left: 200, top: 150, width: 400, height: 250);

        // Assert - Verify position updated
        var readResult = _chartCommands.Read(batch, createResult.ChartName);
        Assert.Equal(200, readResult.Left);
        Assert.Equal(150, readResult.Top);
        Assert.Equal(400, readResult.Width);
        Assert.Equal(250, readResult.Height);
    }

    [Fact]
    public void List_MultipleCharts_ReturnsAll()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create multiple charts
        _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.ColumnClustered, 50, 50, 300, 200, "Chart1");
        _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Pie, 400, 50, 300, 200, "Chart2");
        _chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", ChartType.Line, 50, 300, 300, 200, "Chart3");

        // Act
        var charts = _chartCommands.List(batch);

        // Assert
        Assert.True(charts.Success);
        Assert.Equal(3, charts.Charts.Count);
        Assert.Contains(charts.Charts, c => c.Name == "Chart1" && c.ChartType == ChartType.ColumnClustered);
        Assert.Contains(charts.Charts, c => c.Name == "Chart2" && c.ChartType == ChartType.Pie);
        Assert.Contains(charts.Charts, c => c.Name == "Chart3" && c.ChartType == ChartType.Line);
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
            var result = _chartCommands.CreateFromRange(batch, _sheetName, "A1:C5", chartType, x, 50, 250, 200);
            Assert.Equal(chartType, result.ChartType);
            x += 300;
        }

        // Verify all created
        var charts = _chartCommands.List(batch);
        Assert.True(charts.Success);
        Assert.Equal(chartTypes.Length, charts.Charts.Count);
    }
}
