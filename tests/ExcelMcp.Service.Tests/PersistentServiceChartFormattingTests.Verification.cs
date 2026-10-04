using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceChartFormattingTests
{
    private readonly record struct SeriesState(
        bool HasLabels, bool ShowValue, bool ShowPercentage, int? LabelPosition,
        int? MarkerStyle, int? MarkerSize, int? MarkerFill, int? MarkerBorder);

    private T InspectChart<T>(string chartName, Func<Excel.Chart, T> inspect, string? sheetName = null) =>
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.ChartObject? chartObject = null;
            Excel.Chart? chart = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName ?? _sheetName];
                chartObject = (Excel.ChartObject)sheet.ChartObjects(chartName);
                chart = chartObject.Chart;
                return inspect(chart);
            }
            finally
            {
                ComUtilities.Release(ref chart);
                ComUtilities.Release(ref chartObject);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });

    private void AssertSeriesData(
        string chartName, int seriesIndex, string? expectedName,
        object[] expectedCategories, double[] expectedValues, string? sheetName = null)
    {
        var state = InspectChart(chartName, chart =>
        {
            Excel.SeriesCollection? collection = null;
            Excel.Series? series = null;
            try
            {
                collection = (Excel.SeriesCollection)chart.SeriesCollection();
                series = collection.Item(seriesIndex);
                return (
                    Name: series.Name,
                    Categories: Assert.IsAssignableFrom<Array>((object)series.XValues).Cast<object>().ToArray(),
                    Values: Assert.IsAssignableFrom<Array>((object)series.Values).Cast<object>()
                        .Select(value => Convert.ToDouble(value, CultureInfo.InvariantCulture)).ToArray());
            }
            finally
            {
                ComUtilities.Release(ref series);
                ComUtilities.Release(ref collection);
            }
        }, sheetName);
        if (expectedName != null) { Assert.Equal(expectedName, state.Name); }
        else { Assert.False(string.IsNullOrWhiteSpace(state.Name)); }
        Assert.Equal(expectedCategories.Select(value => value.ToString()), state.Categories.Select(value => value.ToString()));
        Assert.Equal(expectedValues, state.Values);
    }

    private void AssertChartUnchanged(ChartInfoResult before)
    {
        var after = RequireSuccess(_chartCommands.Read(_fixture.BatchToken, before.Name));
        Assert.Equal(before.Name, after.Name);
        Assert.Equal(before.SheetName, after.SheetName);
        Assert.Equal(before.ChartType, after.ChartType);
        Assert.Equal(before.SourceRange, after.SourceRange);
        Assert.Equal(before.Title, after.Title);
        Assert.Equal(before.HasLegend, after.HasLegend);
        Assert.Equal(before.Left, after.Left);
        Assert.Equal(before.Top, after.Top);
        Assert.Equal(before.Width, after.Width);
        Assert.Equal(before.Height, after.Height);
        Assert.Equal(before.Placement, after.Placement);
        Assert.Equal(before.Series.Select(series => series.Name), after.Series.Select(series => series.Name));
    }

    private void AssertFitsRange(string chartName, string sheetName, string rangeAddress)
    {
        var expected = _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                range = sheet.Range[rangeAddress];
                return (Left: Convert.ToDouble(range.Left, CultureInfo.InvariantCulture),
                    Top: Convert.ToDouble(range.Top, CultureInfo.InvariantCulture),
                    Width: Convert.ToDouble(range.Width, CultureInfo.InvariantCulture),
                    Height: Convert.ToDouble(range.Height, CultureInfo.InvariantCulture));
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        var actual = RequireSuccess(_chartCommands.Read(_fixture.BatchToken, chartName));
        Assert.Equal(expected.Left, actual.Left, precision: 2);
        Assert.Equal(expected.Top, actual.Top, precision: 2);
        Assert.Equal(expected.Width, actual.Width, precision: 2);
        Assert.Equal(expected.Height, actual.Height, precision: 2);
        Assert.Equal(_sheetName, actual.SheetName);
    }

    private void AssertAnchorCells(ChartInfoResult actual)
    {
        var native = InspectChart(actual.Name, chart =>
        {
            Excel.ChartObject? chartObject = null;
            Excel.Range? topLeft = null;
            Excel.Range? bottomRight = null;
            try
            {
                chartObject = (Excel.ChartObject)chart.Parent;
                topLeft = chartObject.TopLeftCell;
                bottomRight = chartObject.BottomRightCell;
                return (TopLeft: topLeft.Address, BottomRight: bottomRight.Address);
            }
            finally
            {
                ComUtilities.Release(ref bottomRight);
                ComUtilities.Release(ref topLeft);
                ComUtilities.Release(ref chartObject);
            }
        });
        Assert.Equal(native.TopLeft, actual.TopLeftCell);
        Assert.Equal(native.BottomRight, actual.BottomRightCell);
    }

    private SeriesState ReadSeriesState(string chartName, int seriesIndex = 1) =>
        InspectChart(chartName, chart =>
        {
            Excel.SeriesCollection? collection = null;
            Excel.Series? series = null;
            Excel.DataLabels? labels = null;
            try
            {
                collection = (Excel.SeriesCollection)chart.SeriesCollection();
                series = collection.Item(seriesIndex);
                var supportsPercentage = series.ChartType is
                    Excel.XlChartType.xlPie or Excel.XlChartType.xlDoughnut;
                var supportsMarkers = series.ChartType is
                    Excel.XlChartType.xlLine or Excel.XlChartType.xlLineMarkers
                    or Excel.XlChartType.xlXYScatter;
                if (series.HasDataLabels) { labels = (Excel.DataLabels)series.DataLabels(); }
                return new SeriesState(
                    series.HasDataLabels, labels?.ShowValue ?? false,
                    supportsPercentage && (labels?.ShowPercentage ?? false),
                    labels == null ? null : Convert.ToInt32(labels.Position, CultureInfo.InvariantCulture),
                    supportsMarkers ? Convert.ToInt32(series.MarkerStyle, CultureInfo.InvariantCulture) : null,
                    supportsMarkers ? series.MarkerSize : null,
                    supportsMarkers ? series.MarkerBackgroundColor : null,
                    supportsMarkers ? series.MarkerForegroundColor : null);
            }
            finally
            {
                ComUtilities.Release(ref labels);
                ComUtilities.Release(ref series);
                ComUtilities.Release(ref collection);
            }
        });

    private string ReadAxisTitle(string chartName, ChartAxisType axisType) =>
        InspectChart(chartName, chart =>
        {
            Excel.Axis? axis = null;
            Excel.AxisTitle? title = null;
            try
            {
                var excelAxis = axisType switch
                {
                    ChartAxisType.Category => Excel.XlAxisType.xlCategory,
                    ChartAxisType.Value => Excel.XlAxisType.xlValue,
                    _ => throw new ArgumentOutOfRangeException(nameof(axisType))
                };
                axis = (Excel.Axis)chart.Axes(excelAxis);
                if (!axis.HasTitle) { return ""; }
                title = axis.AxisTitle;
                return title.Text;
            }
            finally
            {
                ComUtilities.Release(ref title);
                ComUtilities.Release(ref axis);
            }
        });

    private int ReadLegendPosition(string chartName) =>
        InspectChart(chartName, chart =>
        {
            Excel.Legend? legend = null;
            try
            {
                legend = chart.Legend;
                return Convert.ToInt32(legend.Position, CultureInfo.InvariantCulture);
            }
            finally { ComUtilities.Release(ref legend); }
        });

    private void AssertNativeAxisScale(AxisScaleResult result) =>
        InspectChart(result.ChartName, chart =>
        {
            Excel.Axis? axis = null;
            try
            {
                axis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlValue);
                Assert.Equal("Value", result.AxisType);
                Assert.Equal(axis.MinimumScaleIsAuto, result.MinimumScaleIsAuto);
                Assert.Equal(axis.MaximumScaleIsAuto, result.MaximumScaleIsAuto);
                Assert.Equal(axis.MajorUnitIsAuto, result.MajorUnitIsAuto);
                Assert.Equal(axis.MinorUnitIsAuto, result.MinorUnitIsAuto);
                Assert.Equal(axis.MinimumScaleIsAuto ? (double?)null : axis.MinimumScale, result.MinimumScale);
                Assert.Equal(axis.MaximumScaleIsAuto ? (double?)null : axis.MaximumScale, result.MaximumScale);
                Assert.Equal(axis.MajorUnitIsAuto ? (double?)null : axis.MajorUnit, result.MajorUnit);
                Assert.Equal(axis.MinorUnitIsAuto ? (double?)null : axis.MinorUnit, result.MinorUnit);
                return 0;
            }
            finally { ComUtilities.Release(ref axis); }
        });

    private void AssertNativeGridlines(GridlinesResult result) =>
        InspectChart(result.ChartName, chart =>
        {
            Excel.Axis? value = null;
            Excel.Axis? category = null;
            try
            {
                value = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlValue);
                category = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlCategory);
                Assert.Equal(value.HasMajorGridlines, result.Gridlines.HasValueMajorGridlines);
                Assert.Equal(value.HasMinorGridlines, result.Gridlines.HasValueMinorGridlines);
                Assert.Equal(category.HasMajorGridlines, result.Gridlines.HasCategoryMajorGridlines);
                Assert.Equal(category.HasMinorGridlines, result.Gridlines.HasCategoryMinorGridlines);
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref category);
                ComUtilities.Release(ref value);
            }
        });

    private ChartInfoResult CreateFormattingGuard()
    {
        var created = RequireSuccess(_chartCommands.CreateFromRange(
            _fixture.BatchToken, _sheetName, "A1:B4", ChartType.Line, 50, 60, 400, 250));
        RequireSuccess(_chartCommands.SetAxisNumberFormat(
            _fixture.BatchToken, created.ChartName, ChartAxisType.Value, "0.00"));
        RequireSuccess(_chartCommands.SetPlacement(_fixture.BatchToken, created.ChartName, 2));
        return RequireSuccess(_chartCommands.Read(_fixture.BatchToken, created.ChartName));
    }

    private void AssertFormattingGuard(ChartInfoResult before)
    {
        AssertChartUnchanged(before);
        Assert.Equal("0.00", _chartCommands.GetAxisNumberFormat(
            _fixture.BatchToken, before.Name, ChartAxisType.Value));
        Assert.Equal(2, ReadChartObjectProperties(before.Name).Placement);
        AssertSeriesData(before.Name, 1, "Series1", ["A", "B", "C"], [10, 15, 20]);
    }
}
