using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceChartFormattingTests
{
    [Fact]
    public void CreateFromPivotTable_RangePivotTable_CreatesPivotChart()
    {
        var batch = _fixture.BatchToken;
        var pivotTableName = $"TestPivot_{Guid.NewGuid():N}";
        var pivotSheetName = _fixture.CreateTestSheet(batch);
        var dashboardSheetName = _fixture.CreateTestSheet(batch);
        CreateRangePivot(
            pivotTableName,
            pivotSheetName,
            "E1:G5",
            new object[,]
            {
                { "Product", "Region", "Sales" },
                { "Widget", "North", 100 },
                { "Widget", "South", 150 },
                { "Gadget", "North", 200 },
                { "Gadget", "South", 250 }
            },
            "Product",
            "Sales");

        var result = _chartCommands.CreateFromPivotTable(
            batch,
            pivotTableName,
            dashboardSheetName,
            ChartType.ColumnClustered,
            300,
            50,
            400,
            300,
            $"PivotChart_{Guid.NewGuid():N}");

        RequireSuccess(result);
        Assert.True(result.IsPivotChart, "Chart should be marked as PivotChart");
        Assert.Equal(pivotTableName, result.LinkedPivotTable);
        Assert.Equal(dashboardSheetName, result.SheetName);
        Assert.Equal(ChartType.ColumnClustered, result.ChartType);

        var chartInfo = RequireSuccess(_chartCommands.Read(batch, result.ChartName));
        Assert.True(chartInfo.IsPivotChart);
        Assert.Equal(pivotTableName, chartInfo.LinkedPivotTable);

        // A single-value PivotChart uses Excel's localized total caption.
        AssertSeriesData(result.ChartName, 1, null, ["Gadget", "Widget"], [450, 250], dashboardSheetName);
        var pivotCommands = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        var added = pivotCommands.AddValueField(batch, pivotTableName, "Region", AggregationFunction.Count, "Verified Count");
        RequireSuccess(added);
        var refreshed = pivotCommands.Refresh(batch, pivotTableName);
        RequireSuccess(refreshed);

        var charts = _chartCommands.List(batch);
        RequireSuccess(charts);
        var linkedChart = Assert.Single(
            charts.Charts,
            chart => chart.Name == result.ChartName);
        Assert.True(linkedChart.IsPivotChart);
        Assert.Equal(pivotTableName, linkedChart.LinkedPivotTable);
        Assert.Equal(2, linkedChart.SeriesCount);
        AssertSeriesData(result.ChartName, 1, "Verified Sales", ["Gadget", "Widget"], [450, 250], dashboardSheetName);
        AssertSeriesData(result.ChartName, 2, "Verified Count", ["Gadget", "Widget"], [2, 2], dashboardSheetName);
    }

    [Fact]
    public void CreateFromPivotTable_UnsupportedPivotChartType_ThrowsWithoutLeavingRegularChart()
    {
        var batch = _fixture.BatchToken;
        var pivotTableName = $"UnsupportedPivot_{Guid.NewGuid():N}";
        var pivotSheetName = _fixture.CreateTestSheet(batch);
        CreateRangePivot(
            pivotTableName,
            pivotSheetName,
            "E10:F13",
            new object[,]
            {
                { "X", "Y" },
                { 1, 2 },
                { 3, 4 },
                { 5, 6 }
            },
            "X",
            "Y");
        var before = RequireSuccess(_chartCommands.List(batch));
        var chartCountBefore = before.Charts.Count(
            chart => chart.SheetName == pivotSheetName);
        var pivotCommands = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        var originalData = pivotCommands.GetData(batch, pivotTableName);
        RequireSuccess(originalData);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _chartCommands.CreateFromPivotTable(
                batch,
                pivotTableName,
                pivotSheetName,
                ChartType.XYScatter,
                300,
                50,
                400,
                300,
                $"UnsupportedPivotChart_{Guid.NewGuid():N}"));

        Assert.Contains(
            "linked PivotChart",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
        var after = RequireSuccess(_chartCommands.List(batch));
        var chartCountAfter = after.Charts.Count(
            chart => chart.SheetName == pivotSheetName);
        Assert.Equal(chartCountBefore, chartCountAfter);
        Assert.Equal(before.Charts.Select(chart => chart.Name), after.Charts.Select(chart => chart.Name));
        var data = pivotCommands.GetData(batch, pivotTableName);
        RequireSuccess(data);
        Assert.Equal(originalData.Values.Count, data.Values.Count);
        for (var index = 0; index < originalData.Values.Count; index++)
        {
            Assert.Equal(originalData.Values[index], data.Values[index]);
        }
    }

    [Fact]
    public void CreateFromPivotTable_DifferentChartTypes_CreatesCorrectType()
    {
        var batch = _fixture.BatchToken;
        var pivotTableName = $"ChartTypePivot_{Guid.NewGuid():N}";
        var pivotSheetName = _fixture.CreateTestSheet(batch);
        CreateRangePivot(
            pivotTableName,
            pivotSheetName,
            "E20:F23",
            new object[,]
            {
                { "Category", "Value" },
                { "A", 10 },
                { "B", 20 },
                { "C", 30 }
            },
            "Category",
            "Value");

        var result = _chartCommands.CreateFromPivotTable(
            batch,
            pivotTableName,
            pivotSheetName,
            ChartType.Pie,
            300,
            50,
            300,
            300);

        Assert.Equal(ChartType.Pie, result.ChartType);
        Assert.True(result.IsPivotChart);
        RequireSuccess(result);
        var read = RequireSuccess(_chartCommands.Read(batch, result.ChartName));
        Assert.Equal(ChartType.Pie, read.ChartType);
        Assert.Equal(pivotTableName, read.LinkedPivotTable);
        AssertSeriesData(result.ChartName, 1, null, ["A", "B", "C"], [10, 20, 30], pivotSheetName);
    }

    private void CreateRangePivot(
        string pivotTableName,
        string pivotSheetName,
        string sourceAddress,
        object[,] values,
        string rowFieldName,
        string dataFieldName)
    {
        var rows = Enumerable.Range(0, values.GetLength(0))
            .Select(row => Enumerable.Range(0, values.GetLength(1))
                .Select(column => (object?)values[row, column]).ToList()).ToList();
        var batch = _fixture.BatchToken;
        RequireSuccess(_commands.SetValues(batch, _sheetName, sourceAddress, rows));
        var pivotCommands = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        var created = pivotCommands.CreateFromRange(batch, _sheetName, sourceAddress,
            pivotSheetName, "A1", pivotTableName);
        RequireSuccess(created);
        var row = pivotCommands.AddRowField(batch, pivotTableName, rowFieldName, null);
        RequireSuccess(row);
        var value = pivotCommands.AddValueField(batch, pivotTableName, dataFieldName, AggregationFunction.Sum, "Verified Sales");
        RequireSuccess(value);
        var refreshed = pivotCommands.Refresh(batch, pivotTableName);
        RequireSuccess(refreshed);
    }
}
