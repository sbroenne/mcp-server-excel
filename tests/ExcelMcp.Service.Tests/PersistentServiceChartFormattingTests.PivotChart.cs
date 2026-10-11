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
    public async Task CreateFromPivotTable_UnsupportedPivotChartType_ReportsActualPartialState()
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
        var supported = RequireSuccess(_chartCommands.CreateFromPivotTable(
            batch, pivotTableName, pivotSheetName, ChartType.ColumnClustered, 750, 50,
            chartName: $"SupportedPivotChart_{Guid.NewGuid():N}"));
        var supportedRead = RequireSuccess(_chartCommands.Read(batch, supported.ChartName));
        Assert.True(supportedRead.IsPivotChart);
        Assert.Equal(pivotTableName, supportedRead.LinkedPivotTable);
        AssertSeriesData(supported.ChartName, 1, null, [1, 3, 5], [2, 4, 6], pivotSheetName);
        var before = RequireSuccess(_chartCommands.List(batch));
        var chartCountBefore = before.Charts.Count(
            chart => chart.SheetName == pivotSheetName);
        var pivotCommands = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        var originalData = pivotCommands.GetData(batch, pivotTableName);
        RequireSuccess(originalData);

        var rejected = await _fixture.SendForFailureAsync("chart.create-from-pivottable", new
        {
            pivotTableName,
            sheetName = pivotSheetName,
            chartType = "XYScatter",
            left = 300,
            top = 50,
            width = 400,
            height = 300,
            chartName = $"UnsupportedPivotChart_{Guid.NewGuid():N}"
        });

        var after = RequireSuccess(_chartCommands.List(batch));
        var retained = after.Charts.Where(chart => chart.SheetName == pivotSheetName)
            .Where(chart => !before.Charts.Any(original =>
                original.SheetName == chart.SheetName && original.Name == chart.Name)).ToList();
        Assert.All(before.Charts, original => Assert.Contains(after.Charts,
            chart => chart.SheetName == original.SheetName && chart.Name == original.Name));
        Assert.DoesNotContain("No chart was kept", rejected.ErrorMessage);
        if (retained.Count == 0)
        {
            // Excel can reject XYScatter in AddChart2, before any chart exists.
            Assert.Equal("ComInterop", rejected.ErrorCategory);
            Assert.Equal("COMException", rejected.ExceptionType);
            Assert.DoesNotContain("failed after Excel created", rejected.ErrorMessage);
            Assert.Equal(chartCountBefore, after.Charts.Count(chart => chart.SheetName == pivotSheetName));
        }
        else
        {
            var chart = Assert.Single(retained);
            Assert.Equal(nameof(InvalidOperationException), rejected.ExceptionType);
            LeftoverObjectAssert.Reported(rejected.ErrorMessage!, "chart", chart.Name, pivotSheetName);
            Assert.True(
                rejected.ErrorMessage!.Contains($"Binding the new chart to PivotTable '{pivotTableName}' failed", StringComparison.Ordinal) ||
                rejected.ErrorMessage.Contains($"Verifying the PivotChart link to '{pivotTableName}' failed", StringComparison.Ordinal),
                rejected.ErrorMessage);
            Assert.Equal(ChartType.XYScatter, chart.ChartType);
            var read = RequireSuccess(_chartCommands.Read(batch, chart.Name));
            Assert.Equal(chart.Name, read.Name);
            Assert.Equal(pivotSheetName, read.SheetName);
            Assert.Equal(300, read.Left);
            Assert.Equal(50, read.Top);
            Assert.Equal(400, read.Width);
            Assert.Equal(300, read.Height);
        }
        var data = pivotCommands.GetData(batch, pivotTableName);
        RequireSuccess(data);
        Assert.Equal(originalData.Values.Count, data.Values.Count);
        for (var index = 0; index < originalData.Values.Count; index++)
        {
            Assert.Equal(originalData.Values[index], data.Values[index]);
        }
        AssertSeriesData(supported.ChartName, 1, null, [1, 3, 5], [2, 4, 6], pivotSheetName);
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
