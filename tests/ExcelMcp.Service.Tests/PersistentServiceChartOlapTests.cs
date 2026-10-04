// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Commands.Slicer;
using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for creating charts from OLAP (Data Model) PivotTables.
/// These tests use <see cref="PersistentServiceDataModelFixture"/> which creates a workbook
/// with Power Pivot Data Model, DAX measures, and OLAP-based PivotTables.
///
/// These tests verify that OLAP chart creation follows Excel's linked PivotChart behavior.
/// </summary>
[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Service")]
[Trait("Feature", "Charts")]
[Trait("RequiresExcel", "true")]
public class PersistentServiceChartOlapTests(
    PersistentServiceDataModelFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceDataModelFixture>
{
    private static readonly string[] ReadbackQuarters = ["Q1", "Q2"];
    private static readonly string[] RegionalCategories = ["East", "North", "South", "West"];
    private static readonly double[] RegionalRevenue = [11500, 10500, 12500, 14500];
    private static readonly double[] RegionalAverage = [5750, 5250, 6250, 7250];
    private static readonly string[] DisambiguationCategories = ["TypeA", "TypeB", "TypeC"];
    private static readonly double[] DisambiguationRevenue = [2500, 5500, 800];
    private static readonly string[] RevenueMeasure = ["[Measures].[TotalRevenue]"];
    private static readonly string[] AverageMeasures = ["[Measures].[TotalRevenue]", "[Measures].[PivotChart Average Revenue]"];
    private static readonly string[] AcrMeasure = ["[Measures].[ACR]"];

    private readonly IPersistentChartCommands _chartCommands =
        fixture.CreateCommands<IPersistentChartCommands>();
    private readonly IPersistentPivotTableCommands _pivotCommands =
        fixture.CreateCommands<IPersistentPivotTableCommands>();
    private readonly IDataModelCommands _dataModelCommands =
        fixture.CreateCommands<IDataModelCommands>();

    [Fact]
    public void ReadAndList_OlapColumnMembers_ReportPlottedSeriesAndFollowFilters()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        var pivotName = $"SeriesPivot_{Guid.NewGuid():N}";
        var pivot = _pivotCommands.CreateFromDataModel(
            batch, "RegionalSalesTable", sheet, "A1", pivotName);
        Assert.True(pivot.Success, pivot.ErrorMessage);
        Assert.True(_pivotCommands.AddRowField(
            batch, pivotName, "[RegionalSalesTable].[Quarter]", null).Success);
        Assert.True(_pivotCommands.AddColumnField(
            batch, pivotName, "[RegionalSalesTable].[Region]", null).Success);
        Assert.True(_pivotCommands.AddValueField(
            batch, pivotName, "[Measures].[TotalRevenue]",
            AggregationFunction.Sum, "Revenue").Success);

        var slicers = _fixture.CreateCommands<ISlicerCommands>();
        var slicerName = $"SeriesFilter_{Guid.NewGuid():N}";
        var slicer = slicers.CreateSlicer(
            batch, pivotName, "[RegionalSalesTable].[Region]",
            slicerName, sheet, "J1");
        Assert.True(slicer.Success, slicer.ErrorMessage);
        var members = slicer.AvailableItems.Order(StringComparer.Ordinal).Take(3).ToList();
        Assert.Equal(3, members.Count);
        Assert.True(slicers.SetSlicerSelection(batch, slicerName, members).Success);

        var created = _chartCommands.CreateFromPivotTable(
            batch, pivotName, sheet, ChartType.Line, chartName: $"SeriesChart_{Guid.NewGuid():N}");
        Assert.True(created.Success, created.ErrorMessage);
        AssertSeries(members);

        Assert.True(slicers.SetSlicerSelection(batch, slicerName, [members[0]]).Success);
        AssertSeries([members[0]]);
        Assert.True(slicers.SetSlicerSelection(batch, slicerName, []).Success);
        Assert.True(_pivotCommands.Refresh(batch, pivotName, null).Success);
        AssertSeries(slicer.AvailableItems.Order(StringComparer.Ordinal).ToList());

        void AssertSeries(List<string> expected)
        {
            var listed = _chartCommands.List(batch);
            Assert.True(listed.Success, listed.ErrorMessage);
            var chart = Assert.Single(listed.Charts, c => c.Name == created.ChartName);
            Assert.True(chart.IsPivotChart);
            Assert.Equal(pivotName, chart.LinkedPivotTable);
            Assert.Equal(expected.Count, chart.SeriesCount);

            var read = _chartCommands.Read(batch, created.ChartName);
            Assert.True(read.Success, read.ErrorMessage);
            Assert.True(read.IsPivotChart);
            Assert.Equal(pivotName, read.LinkedPivotTable);
            Assert.Equal(expected.Count, read.Series.Count);
            Assert.Equal(expected, read.Series.Select(s => s.Name).Order(StringComparer.Ordinal).ToList());
            var expectedValues = new Dictionary<string, double[]>(StringComparer.Ordinal)
            {
                ["East"] = [5500, 6000],
                ["North"] = [5000, 5500],
                ["South"] = [6000, 6500],
                ["West"] = [7000, 7500]
            };
            foreach (var series in read.Series)
            {
                Assert.Equal(ReadbackQuarters, series.Categories.Select(value => value?.ToString()));
                Assert.Equal(expectedValues[series.Name], series.Values.Select(value =>
                    Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture)));
                Assert.Equal(string.Empty, series.ValuesRange);
                Assert.Null(series.CategoryRange);
            }
        }
    }

    [Fact]
    public void CreateFromPivotTable_OlapDataModelPivot_CreatesPivotChart()
    {
        // Arrange - Use the Data Model PivotTable from fixture
        // The fixture creates "DataModelPivot" PivotTable on sheet "ModelData"
        var (pivotTableName, sheetName) = PrepareRegionalPivot();

        // Act
        var batch = _fixture.BatchToken;
        var result = _chartCommands.CreateFromPivotTable(
            batch,
            pivotTableName,
            sheetName,
            ChartType.ColumnClustered,
            300,
            200,
            400,
            300,
            "OlapChart1");

        // Assert
        RequireSuccess(result);
        Assert.True(result.IsPivotChart, "Chart should be marked as PivotChart");
        Assert.Equal(pivotTableName, result.LinkedPivotTable);
        Assert.Equal(sheetName, result.SheetName);
        Assert.Equal(ChartType.ColumnClustered, result.ChartType);
        Assert.NotNull(result.ChartName);

        // Verify Excel reports a real PivotChart linked to the source Data Model PivotTable.
        var chartInfo = _chartCommands.Read(batch, result.ChartName);
        RequireSuccess(chartInfo);
        Assert.True(chartInfo.IsPivotChart);
        Assert.Equal(pivotTableName, chartInfo.LinkedPivotTable);
        AssertRegionalSeries(result.ChartName, pivotTableName, ChartType.ColumnClustered);

        // Verify the live link follows OLAP PivotTable field changes.
        var dataModelCommands = _dataModelCommands;
        RequireSuccess(dataModelCommands.CreateMeasure(
            batch,
            "RegionalSalesTable",
            "PivotChart Average Revenue",
            "AVERAGE('RegionalSalesTable'[Sales])",
            formatType: "Decimal"));
        _fixture.RegisterDataModelMeasureForCleanup(
            "PivotChart Average Revenue");

        var pivotCommands = _pivotCommands;
        RequireSuccess(pivotCommands.Refresh(batch, pivotTableName, null));
        RequireSuccess(pivotCommands.AddValueField(
            batch,
            pivotTableName,
            "[Measures].[PivotChart Average Revenue]",
            AggregationFunction.Average,
            "Average Revenue"));
        RequireSuccess(pivotCommands.Refresh(batch, pivotTableName, null));

        // Verify the chart still resolves through PivotLayout and exposes both value fields.
        var charts = _chartCommands.List(batch);
        RequireSuccess(charts);
        var linkedChart = Assert.Single(charts.Charts, c => c.Name == result.ChartName);
        Assert.True(linkedChart.IsPivotChart);
        Assert.Equal(pivotTableName, linkedChart.LinkedPivotTable);
        Assert.Equal(2, linkedChart.SeriesCount);
        var read = RequireSuccess(_chartCommands.Read(batch, result.ChartName));
        Assert.Equal(2, read.Series.Count);
        Assert.Equal(AssertNativeSeriesNames(read.SheetName, result.ChartName, pivotTableName, AverageMeasures),
            read.Series.Select(s => s.Name));
        var average = Assert.Single(read.Series, s => s.Values.Select(Convert.ToDouble).SequenceEqual(RegionalAverage));
        Assert.Equal(RegionalCategories, average.Categories.Select(c => c?.ToString()));
        Assert.Equal(RegionalAverage, average.Values.Select(Convert.ToDouble));
    }

    [Fact]
    public void CreateFromPivotTable_OlapPivotWithDaxMeasures_CreatesPivotChart()
    {
        // Arrange - Use the Data Model PivotTable that includes DAX measures
        // DataModelPivot has measures like "Total Sales", "Total Revenue", etc.
        var (pivotTableName, sheetName) = PrepareRegionalPivot();

        // Act
        var batch = _fixture.BatchToken;
        var result = _chartCommands.CreateFromPivotTable(
            batch,
            pivotTableName,
            sheetName,
            ChartType.Pie,
            300,
            400,
            350,
            350,
            "OlapPieChart");

        // Assert
        RequireSuccess(result);
        Assert.True(result.IsPivotChart);
        Assert.Equal(ChartType.Pie, result.ChartType);

        // Verify chart was created
        var chartInfo = _chartCommands.Read(batch, result.ChartName);
        RequireSuccess(chartInfo);
        Assert.Equal("OlapPieChart", chartInfo.Name);
        Assert.Equal(sheetName, chartInfo.SheetName);
        AssertRegionalSeries(result.ChartName, pivotTableName, ChartType.Pie);
    }

    [Fact]
    public void CreateFromPivotTable_OlapPivot_BarChart_CreatesCorrectType()
    {
        // Arrange
        var (pivotTableName, sheetName) = PrepareRegionalPivot();

        // Act
        var batch = _fixture.BatchToken;
        var result = _chartCommands.CreateFromPivotTable(
            batch,
            pivotTableName,
            sheetName,
            ChartType.BarClustered,
            50,
            500,
            400,
            300,
            "OlapBarChart");

        // Assert
        RequireSuccess(result);
        Assert.True(result.IsPivotChart);
        Assert.Equal(ChartType.BarClustered, result.ChartType);

        // Verify via Read
        var chartInfo = _chartCommands.Read(batch, result.ChartName);
        RequireSuccess(chartInfo);
        Assert.Equal(ChartType.BarClustered, chartInfo.ChartType);
        AssertRegionalSeries(result.ChartName, pivotTableName, ChartType.BarClustered);
    }

    [Fact]
    public void CreateFromPivotTable_OlapPivot_LineChart_CreatesCorrectType()
    {
        // Arrange
        var (pivotTableName, sheetName) = PrepareRegionalPivot();

        // Act
        var batch = _fixture.BatchToken;
        var result = _chartCommands.CreateFromPivotTable(
            batch,
            pivotTableName,
            sheetName,
            ChartType.Line,
            50,
            700,
            400,
            300,
            "OlapLineChart");

        // Assert
        RequireSuccess(result);
        Assert.True(result.IsPivotChart);
        Assert.Equal(ChartType.Line, result.ChartType);

        // Verify chart appears in list
        var charts = _chartCommands.List(batch);
        RequireSuccess(charts);
        Assert.Contains(charts.Charts, c => c.Name == "OlapLineChart" && c.ChartType == ChartType.Line);
        AssertRegionalSeries(result.ChartName, pivotTableName, ChartType.Line);
    }

    [Fact]
    public void CreateFromPivotTable_DisambiguationTestPivot_CreatesPivotChart()
    {
        // Arrange - Use the second OLAP PivotTable created by fixture
        // "DisambiguationTest" is on sheet "DisambiguationPivot"
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var pivotTableName = $"ChartMeasure_{Guid.NewGuid():N}";
        RequireSuccess(_pivotCommands.CreateFromDataModel(_fixture.BatchToken,
            "DisambiguationTable", sheetName, "A1", pivotTableName));
        RequireSuccess(_pivotCommands.AddRowField(_fixture.BatchToken, pivotTableName,
            "[DisambiguationTable].[Category]", null));
        RequireSuccess(_pivotCommands.AddValueField(_fixture.BatchToken, pivotTableName,
            "[Measures].[ACR]", AggregationFunction.Sum, null));

        // Act
        var batch = _fixture.BatchToken;
        var result = _chartCommands.CreateFromPivotTable(
            batch,
            pivotTableName,
            sheetName,
            ChartType.ColumnClustered,
            50,
            200,
            400,
            300,
            "DisambiguationChart");

        // Assert
        RequireSuccess(result);
        Assert.True(result.IsPivotChart);
        Assert.Equal(pivotTableName, result.LinkedPivotTable);
        Assert.Equal(sheetName, result.SheetName);

        // Verify chart exists
        var charts = _chartCommands.List(batch);
        RequireSuccess(charts);
        Assert.Contains(charts.Charts, c => c.Name == "DisambiguationChart");
        var read = RequireSuccess(_chartCommands.Read(batch, result.ChartName));
        Assert.True(read.IsPivotChart);
        Assert.Equal(pivotTableName, read.LinkedPivotTable);
        var series = Assert.Single(read.Series);
        Assert.Equal("Total", series.Name);
        Assert.Equal(AssertNativeSeriesNames(read.SheetName, result.ChartName, pivotTableName, AcrMeasure),
            read.Series.Select(s => s.Name));
        Assert.Equal(DisambiguationCategories, series.Categories.Select(c => c?.ToString()));
        Assert.Equal(DisambiguationRevenue, series.Values.Select(Convert.ToDouble));
    }

    [Fact]
    public void CreateFromPivotTable_OlapPivot_CustomPositionAndSize_AppliesCorrectly()
    {
        // Arrange
        var (pivotTableName, sheetName) = PrepareRegionalPivot();

        double expectedLeft = 150;
        double expectedTop = 100;
        double expectedWidth = 500;
        double expectedHeight = 400;

        // Act
        var batch = _fixture.BatchToken;
        var result = _chartCommands.CreateFromPivotTable(
            batch,
            pivotTableName,
            sheetName,
            ChartType.ColumnClustered,
            expectedLeft,
            expectedTop,
            expectedWidth,
            expectedHeight,
            "PositionedOlapChart");

        // Assert
        RequireSuccess(result);
        Assert.Equal(expectedLeft, result.Left);
        Assert.Equal(expectedTop, result.Top);
        Assert.Equal(expectedWidth, result.Width);
        Assert.Equal(expectedHeight, result.Height);

        // Verify via Read
        var chartInfo = _chartCommands.Read(batch, result.ChartName);
        RequireSuccess(chartInfo);
        Assert.Equal(expectedLeft, chartInfo.Left);
        Assert.Equal(expectedTop, chartInfo.Top);
        Assert.Equal(expectedWidth, chartInfo.Width);
        Assert.Equal(expectedHeight, chartInfo.Height);
        AssertRegionalSeries(result.ChartName, pivotTableName, ChartType.ColumnClustered);
    }

    private (string Pivot, string Sheet) PrepareRegionalPivot()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        var name = $"ChartPivot_{Guid.NewGuid():N}";
        RequireSuccess(_pivotCommands.CreateFromDataModel(_fixture.BatchToken,
            "RegionalSalesTable", sheet, "A1", name));
        RequireSuccess(_pivotCommands.AddRowField(_fixture.BatchToken, name,
            "[RegionalSalesTable].[Region]", null));
        RequireSuccess(_pivotCommands.AddValueField(_fixture.BatchToken, name,
            "[Measures].[TotalRevenue]", AggregationFunction.Sum, null));
        RequireSuccess(_pivotCommands.Refresh(_fixture.BatchToken, name, null));
        return (name, sheet);
    }

    private void AssertRegionalSeries(string chartName, string pivotName, ChartType type)
    {
        var read = RequireSuccess(_chartCommands.Read(_fixture.BatchToken, chartName));
        Assert.True(read.IsPivotChart);
        Assert.Equal(pivotName, read.LinkedPivotTable);
        Assert.Equal(type, read.ChartType);
        var series = Assert.Single(read.Series);
        // Excel uses "Total" for a single-value-field PivotChart series.
        Assert.Equal("Total", series.Name);
        Assert.Equal(AssertNativeSeriesNames(read.SheetName, chartName, pivotName, RevenueMeasure),
            read.Series.Select(s => s.Name));
        Assert.Equal(RegionalCategories, series.Categories.Select(c => c?.ToString()));
        Assert.Equal(RegionalRevenue, series.Values.Select(Convert.ToDouble));
    }

    private string[] AssertNativeSeriesNames(string sheetName, string chartName, string pivotName,
        string[] expectedMeasures) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.ChartObjects? charts = null;
            Excel.ChartObject? chartObject = null;
            Excel.Chart? chart = null;
            Excel.PivotLayout? layout = null;
            Excel.PivotTable? pivot = null;
            Excel.PivotFields? fields = null;
            Excel.SeriesCollection? seriesCollection = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                charts = (Excel.ChartObjects)sheet.ChartObjects();
                chartObject = (Excel.ChartObject)charts.Item(chartName);
                chart = chartObject.Chart;
                layout = chart.PivotLayout;
                Assert.NotNull(layout);
                pivot = layout.PivotTable;
                Assert.Equal(pivotName, pivot.Name);
                fields = (Excel.PivotFields)pivot.DataFields;
                var measures = new List<string>();
                for (int index = 1; index <= fields.Count; index++)
                {
                    Excel.PivotField? field = null;
                    Excel.CubeField? cube = null;
                    try
                    {
                        field = fields.Item(index);
                        cube = field.CubeField;
                        measures.Add(cube.Name);
                    }
                    finally
                    {
                        ComUtilities.Release(ref cube);
                        ComUtilities.Release(ref field);
                    }
                }
                Assert.Equal(expectedMeasures, measures);
                seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection();
                var names = new List<string>();
                for (int index = 1; index <= seriesCollection.Count; index++)
                {
                    Excel.Series? series = null;
                    try
                    {
                        series = seriesCollection.Item(index);
                        names.Add(series.Name);
                    }
                    finally
                    {
                        ComUtilities.Release(ref series);
                    }
                }
                return names.ToArray();
            }
            finally
            {
                ComUtilities.Release(ref seriesCollection);
                ComUtilities.Release(ref fields);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref layout);
                ComUtilities.Release(ref chart);
                ComUtilities.Release(ref chartObject);
                ComUtilities.Release(ref charts);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
}
