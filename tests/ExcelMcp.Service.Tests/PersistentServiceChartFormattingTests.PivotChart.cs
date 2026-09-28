using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

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

        Assert.True(result.IsPivotChart, "Chart should be marked as PivotChart");
        Assert.Equal(pivotTableName, result.LinkedPivotTable);
        Assert.Equal(dashboardSheetName, result.SheetName);
        Assert.Equal(ChartType.ColumnClustered, result.ChartType);

        var chartInfo = _chartCommands.Read(batch, result.ChartName);
        Assert.True(chartInfo.IsPivotChart);
        Assert.Equal(pivotTableName, chartInfo.LinkedPivotTable);

        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic? pivotSheet = null;
            dynamic? pivotTable = null;
            dynamic? secondDataField = null;
            try
            {
                pivotSheet = ctx.Book.Worksheets[pivotSheetName];
                pivotTable = pivotSheet.PivotTables(pivotTableName);
                secondDataField = pivotTable.PivotFields("Region");
                secondDataField.Orientation =
                    (int)Excel.XlPivotFieldOrientation.xlDataField;
                secondDataField.Function =
                    (int)Excel.XlConsolidationFunction.xlCount;
                pivotTable.RefreshTable();
            }
            finally
            {
                ComUtilities.Release(ref secondDataField);
                ComUtilities.Release(ref pivotTable);
                ComUtilities.Release(ref pivotSheet);
            }
        });

        var charts = _chartCommands.List(batch);
        Assert.True(charts.Success);
        var linkedChart = Assert.Single(
            charts.Charts,
            chart => chart.Name == result.ChartName);
        Assert.True(linkedChart.IsPivotChart);
        Assert.Equal(pivotTableName, linkedChart.LinkedPivotTable);
        Assert.Equal(2, linkedChart.SeriesCount);
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
        var chartCountBefore = _chartCommands.List(batch).Charts.Count(
            chart => chart.SheetName == pivotSheetName);

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
        var chartCountAfter = _chartCommands.List(batch).Charts.Count(
            chart => chart.SheetName == pivotSheetName);
        Assert.Equal(chartCountBefore, chartCountAfter);
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
    }

    private void CreateRangePivot(
        string pivotTableName,
        string pivotSheetName,
        string sourceAddress,
        object[,] values,
        string rowFieldName,
        string dataFieldName) =>
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic? sourceSheet = null;
            dynamic? sourceRange = null;
            dynamic? pivotCaches = null;
            dynamic? pivotCache = null;
            dynamic? pivotSheet = null;
            dynamic? pivotDestination = null;
            dynamic? pivotTable = null;
            dynamic? rowField = null;
            dynamic? dataField = null;
            try
            {
                sourceSheet = ctx.Book.Worksheets[_sheetName];
                sourceRange = sourceSheet.Range[sourceAddress];
                sourceRange.Value2 = values;
                pivotCaches = ctx.Book.PivotCaches();
                pivotCache = pivotCaches.Create(
                    Excel.XlPivotTableSourceType.xlDatabase,
                    sourceRange);
                pivotSheet = ctx.Book.Worksheets[pivotSheetName];
                pivotDestination = pivotSheet.Range["A1"];
                pivotTable = pivotCache.CreatePivotTable(
                    pivotDestination,
                    pivotTableName);
                rowField = pivotTable.PivotFields(rowFieldName);
                rowField.Orientation =
                    (int)Excel.XlPivotFieldOrientation.xlRowField;
                dataField = pivotTable.PivotFields(dataFieldName);
                dataField.Orientation =
                    (int)Excel.XlPivotFieldOrientation.xlDataField;
            }
            finally
            {
                ComUtilities.Release(ref dataField);
                ComUtilities.Release(ref rowField);
                ComUtilities.Release(ref pivotTable);
                ComUtilities.Release(ref pivotDestination);
                ComUtilities.Release(ref pivotSheet);
                ComUtilities.Release(ref pivotCache);
                ComUtilities.Release(ref pivotCaches);
                ComUtilities.Release(ref sourceRange);
                ComUtilities.Release(ref sourceSheet);
            }
        });
}
