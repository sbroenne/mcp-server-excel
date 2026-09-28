using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceChartFormattingTests
{
    [Fact]
    public void SetSeriesChartType_CreatesColumnLineComboChart()
    {
        var batch = _fixture.BatchToken;
        var createResult = _chartCommands.CreateFromRange(
            batch,
            _sheetName,
            "A1:C6",
            ChartType.ColumnClustered,
            chartName: $"ComboChart_{Guid.NewGuid():N}");

        var result = _chartCommands.SetSeriesChartType(
            batch,
            createResult.ChartName,
            seriesIndex: 2,
            ChartType.LineMarkers);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(
            ChartType.LineMarkers,
            ReadSeriesChartType(createResult.ChartName, 2));
    }

    [Fact]
    public void SetPlacement_ConfiguresEmbeddedChartObjectProperties()
    {
        var batch = _fixture.BatchToken;
        var createResult = _chartCommands.CreateFromRange(
            batch,
            _sheetName,
            "A1:B6",
            ChartType.ColumnClustered,
            chartName: $"ObjectProperties_{Guid.NewGuid():N}");

        var result = _chartCommands.SetPlacement(
            batch,
            createResult.ChartName,
            placement: 2,
            printObject: false,
            locked: false,
            roundedCorners: true);

        Assert.True(result.Success, result.ErrorMessage);
        var properties = ReadChartObjectProperties(createResult.ChartName);
        Assert.Equal(2, properties.Placement);
        Assert.False(properties.PrintObject);
        Assert.False(properties.Locked);
        Assert.True(properties.RoundedCorners);
    }

    [Fact]
    public void SetAreaFormat_AppliesChartAreaFillAndBorder()
    {
        var batch = _fixture.BatchToken;
        var createResult = _chartCommands.CreateFromRange(
            batch,
            _sheetName,
            "A1:B6",
            ChartType.ColumnClustered,
            chartName: $"AreaFormat_{Guid.NewGuid():N}");

        var result = _chartCommands.SetAreaFormat(
            batch,
            createResult.ChartName,
            ChartAreaTarget.Chart,
            fillColor: "#FF0000",
            fillTransparency: 0.25,
            lineColor: "#0000FF",
            lineWeight: 2.5);

        Assert.True(result.Success, result.ErrorMessage);
        var format = ReadChartAreaFormat(createResult.ChartName);
        Assert.Equal(0x0000FF, format.FillColor);
        Assert.Equal(0.25f, format.FillTransparency, precision: 2);
        Assert.Equal(0xFF0000, format.LineColor);
        Assert.Equal(2.5f, format.LineWeight, precision: 2);
    }

    [Fact]
    public void SetSeriesFormat_AppliesSeriesFillAndLineMaterialFormatting()
    {
        var batch = _fixture.BatchToken;
        var createResult = _chartCommands.CreateFromRange(
            batch,
            _sheetName,
            "A1:B6",
            ChartType.ColumnClustered,
            chartName: $"SeriesMaterial_{Guid.NewGuid():N}");

        var result = _chartCommands.SetSeriesFormat(
            batch,
            createResult.ChartName,
            seriesIndex: 1,
            fillColor: "#00FF00",
            fillTransparency: 0.4,
            lineColor: "#FF00FF",
            lineWeight: 3);

        Assert.True(result.Success, result.ErrorMessage);
        var format = ReadSeriesFormat(createResult.ChartName, 1);
        Assert.Equal(0x00FF00, format.FillColor);
        Assert.Equal(0.4f, format.FillTransparency, precision: 2);
        Assert.Equal(0xFF00FF, format.LineColor);
        Assert.Equal(3f, format.LineWeight, precision: 2);
    }

    private ChartType ReadSeriesChartType(string chartName, int seriesIndex) =>
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.ChartObject? chartObject = null;
            Excel.Chart? chart = null;
            Excel.SeriesCollection? seriesCollection = null;
            Excel.Series? series = null;
            try
            {
                sheet = (Excel.Worksheet)ctx.Book.Worksheets[_sheetName];
                chartObject = (Excel.ChartObject)sheet.ChartObjects(chartName);
                chart = chartObject.Chart;
                seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection();
                series = seriesCollection.Item(seriesIndex);
                return (ChartType)series.ChartType;
            }
            finally
            {
                ComUtilities.Release(ref series);
                ComUtilities.Release(ref seriesCollection);
                ComUtilities.Release(ref chart);
                ComUtilities.Release(ref chartObject);
                ComUtilities.Release(ref sheet);
            }
        });

    private (int Placement, bool PrintObject, bool Locked, bool RoundedCorners)
        ReadChartObjectProperties(string chartName) =>
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.ChartObject? chartObject = null;
            try
            {
                sheet = (Excel.Worksheet)ctx.Book.Worksheets[_sheetName];
                chartObject = (Excel.ChartObject)sheet.ChartObjects(chartName);
                return (
                    Convert.ToInt32(
                        (object)chartObject.Placement,
                        CultureInfo.InvariantCulture),
                    chartObject.PrintObject,
                    chartObject.Locked,
                    chartObject.RoundedCorners);
            }
            finally
            {
                ComUtilities.Release(ref chartObject);
                ComUtilities.Release(ref sheet);
            }
        });

    private (int FillColor, float FillTransparency, int LineColor, float LineWeight)
        ReadChartAreaFormat(string chartName) =>
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.ChartObject? chartObject = null;
            Excel.Chart? chart = null;
            Excel.ChartArea? chartArea = null;
            dynamic? chartFormat = null;
            dynamic? fill = null;
            dynamic? line = null;
            dynamic? fillColor = null;
            dynamic? lineColor = null;
            try
            {
                sheet = (Excel.Worksheet)ctx.Book.Worksheets[_sheetName];
                chartObject = (Excel.ChartObject)sheet.ChartObjects(chartName);
                chart = chartObject.Chart;
                chartArea = chart.ChartArea;
                chartFormat = chartArea.Format;
                fill = chartFormat.Fill;
                line = chartFormat.Line;
                fillColor = fill.ForeColor;
                lineColor = line.ForeColor;
                return (
                    (int)fillColor.RGB,
                    (float)fill.Transparency,
                    (int)lineColor.RGB,
                    (float)line.Weight);
            }
            finally
            {
                ComUtilities.Release(ref lineColor);
                ComUtilities.Release(ref fillColor);
                ComUtilities.Release(ref line);
                ComUtilities.Release(ref fill);
                ComUtilities.Release(ref chartFormat);
                ComUtilities.Release(ref chartArea);
                ComUtilities.Release(ref chart);
                ComUtilities.Release(ref chartObject);
                ComUtilities.Release(ref sheet);
            }
        });

    private (int FillColor, float FillTransparency, int LineColor, float LineWeight)
        ReadSeriesFormat(string chartName, int seriesIndex) =>
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.ChartObject? chartObject = null;
            Excel.Chart? chart = null;
            Excel.SeriesCollection? seriesCollection = null;
            Excel.Series? series = null;
            dynamic? chartFormat = null;
            dynamic? fill = null;
            dynamic? line = null;
            dynamic? fillColor = null;
            dynamic? lineColor = null;
            try
            {
                sheet = (Excel.Worksheet)ctx.Book.Worksheets[_sheetName];
                chartObject = (Excel.ChartObject)sheet.ChartObjects(chartName);
                chart = chartObject.Chart;
                seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection();
                series = seriesCollection.Item(seriesIndex);
                chartFormat = series.Format;
                fill = chartFormat.Fill;
                line = chartFormat.Line;
                fillColor = fill.ForeColor;
                lineColor = line.ForeColor;
                return (
                    (int)fillColor.RGB,
                    (float)fill.Transparency,
                    (int)lineColor.RGB,
                    (float)line.Weight);
            }
            finally
            {
                ComUtilities.Release(ref lineColor);
                ComUtilities.Release(ref fillColor);
                ComUtilities.Release(ref line);
                ComUtilities.Release(ref fill);
                ComUtilities.Release(ref chartFormat);
                ComUtilities.Release(ref series);
                ComUtilities.Release(ref seriesCollection);
                ComUtilities.Release(ref chart);
                ComUtilities.Release(ref chartObject);
                ComUtilities.Release(ref sheet);
            }
        });
}
