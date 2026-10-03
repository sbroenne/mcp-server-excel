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
        RequireSuccess(createResult);
        Assert.Equal(ChartType.ColumnClustered, ReadSeriesChartType(createResult.ChartName, 1));
        Assert.Equal(ChartType.ColumnClustered, ReadSeriesChartType(createResult.ChartName, 2));

        var result = _chartCommands.SetSeriesChartType(
            batch,
            createResult.ChartName,
            seriesIndex: 2,
            ChartType.LineMarkers);

        RequireSuccess(result);
        Assert.Equal(
            ChartType.LineMarkers,
            ReadSeriesChartType(createResult.ChartName, 2));
        Assert.Equal(ChartType.ColumnClustered, ReadSeriesChartType(createResult.ChartName, 1));
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
        RequireSuccess(createResult);
        RequireSuccess(_chartCommands.SetPlacement(batch, createResult.ChartName, 1,
            printObject: true, locked: true, roundedCorners: false));
        Assert.Equal((1, true, true, false), ReadChartObjectProperties(createResult.ChartName));

        var result = _chartCommands.SetPlacement(
            batch,
            createResult.ChartName,
            placement: 2,
            printObject: false,
            locked: false,
            roundedCorners: true);

        RequireSuccess(result);
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
        RequireSuccess(createResult);

        var result = _chartCommands.SetAreaFormat(
            batch,
            createResult.ChartName,
            ChartAreaTarget.Chart,
            fillColor: "#FF0000",
            fillTransparency: 0.25,
            lineColor: "#0000FF",
            lineWeight: 2.5);

        RequireSuccess(result);
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
        RequireSuccess(createResult);

        var result = _chartCommands.SetSeriesFormat(
            batch,
            createResult.ChartName,
            seriesIndex: 1,
            fillColor: "#00FF00",
            fillTransparency: 0.4,
            lineColor: "#FF00FF",
            lineWeight: 3);

        RequireSuccess(result);
        var format = ReadSeriesFormat(createResult.ChartName, 1);
        Assert.Equal(0x00FF00, format.FillColor);
        Assert.Equal(0.4f, format.FillTransparency, precision: 2);
        Assert.Equal(0xFF00FF, format.LineColor);
        Assert.Equal(3f, format.LineWeight, precision: 2);
    }

    private ChartType ReadSeriesChartType(string chartName, int seriesIndex) =>
        InspectChart(chartName, chart =>
        {
            Excel.SeriesCollection? seriesCollection = null;
            Excel.Series? series = null;
            try
            {
                seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection();
                series = seriesCollection.Item(seriesIndex);
                return (ChartType)series.ChartType;
            }
            finally
            {
                ComUtilities.Release(ref series);
                ComUtilities.Release(ref seriesCollection);
            }
        });

    private (int Placement, bool PrintObject, bool Locked, bool RoundedCorners)
        ReadChartObjectProperties(string chartName) =>
        InspectChart(chartName, chart =>
        {
            Excel.ChartObject? chartObject = null;
            try
            {
                chartObject = (Excel.ChartObject)chart.Parent;
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
            }
        });

    private (int FillColor, float FillTransparency, int LineColor, float LineWeight)
        ReadChartAreaFormat(string chartName) =>
        InspectChart(chartName, chart =>
        {
            Excel.ChartArea? chartArea = null;
            // Office-core formatting objects are deliberately late-bound.
            dynamic? chartFormat = null;
            dynamic? fill = null;
            dynamic? line = null;
            dynamic? fillColor = null;
            dynamic? lineColor = null;
            try
            {
                chartArea = chart.ChartArea;
                chartFormat = chartArea.Format;
                fill = chartFormat.Fill;
                line = chartFormat.Line;
                fillColor = fill.ForeColor;
                lineColor = line.ForeColor;
                return (
                    Convert.ToInt32((object)fillColor.RGB, CultureInfo.InvariantCulture),
                    Convert.ToSingle((object)fill.Transparency, CultureInfo.InvariantCulture),
                    Convert.ToInt32((object)lineColor.RGB, CultureInfo.InvariantCulture),
                    Convert.ToSingle((object)line.Weight, CultureInfo.InvariantCulture));
            }
            finally
            {
                ComUtilities.Release(ref lineColor);
                ComUtilities.Release(ref fillColor);
                ComUtilities.Release(ref line);
                ComUtilities.Release(ref fill);
                ComUtilities.Release(ref chartFormat);
                ComUtilities.Release(ref chartArea);
            }
        });

    private (int FillColor, float FillTransparency, int LineColor, float LineWeight)
        ReadSeriesFormat(string chartName, int seriesIndex) =>
        InspectChart(chartName, chart =>
        {
            Excel.SeriesCollection? seriesCollection = null;
            Excel.Series? series = null;
            // Office-core formatting objects are deliberately late-bound.
            dynamic? chartFormat = null;
            dynamic? fill = null;
            dynamic? line = null;
            dynamic? fillColor = null;
            dynamic? lineColor = null;
            try
            {
                seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection();
                series = seriesCollection.Item(seriesIndex);
                chartFormat = series.Format;
                fill = chartFormat.Fill;
                line = chartFormat.Line;
                fillColor = fill.ForeColor;
                lineColor = line.ForeColor;
                return (
                    Convert.ToInt32((object)fillColor.RGB, CultureInfo.InvariantCulture),
                    Convert.ToSingle((object)fill.Transparency, CultureInfo.InvariantCulture),
                    Convert.ToInt32((object)lineColor.RGB, CultureInfo.InvariantCulture),
                    Convert.ToSingle((object)line.Weight, CultureInfo.InvariantCulture));
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
            }
        });
}
