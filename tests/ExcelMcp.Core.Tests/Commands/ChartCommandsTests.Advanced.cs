using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

public partial class ChartCommandsTests
{
    [Fact]
    public void SetSeriesChartType_CreatesColumnLineComboChart()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);
        var createResult = _commands.CreateFromRange(
            batch, "Sheet1", "A1:C6", ChartType.ColumnClustered, chartName: "ComboChart");

        var result = _commands.SetSeriesChartType(
            batch, createResult.ChartName, seriesIndex: 2, ChartType.LineMarkers);

        Assert.True(result.Success, result.ErrorMessage);
        var actualType = ReadSeriesChartType(batch, createResult.ChartName, 2);
        Assert.Equal(ChartType.LineMarkers, actualType);
    }

    [Fact]
    public void PlotOptions_SetAndGet_RoundTripsChartBehavior()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);
        var createResult = _commands.CreateFromRange(
            batch, "Sheet1", "A1:C6", ChartType.Line, chartName: "PlotOptionsChart");

        var setResult = _commands.SetPlotOptions(
            batch,
            createResult.ChartName,
            plotBy: ChartPlotBy.Rows,
            displayBlanksAs: ChartDisplayBlanksAs.Zero,
            plotVisibleOnly: false);

        Assert.True(setResult.Success, setResult.ErrorMessage);
        var getResult = _commands.GetPlotOptions(batch, createResult.ChartName);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.Equal(ChartPlotBy.Rows, getResult.PlotBy);
        Assert.Equal(ChartDisplayBlanksAs.Zero, getResult.DisplayBlanksAs);
        Assert.False(getResult.PlotVisibleOnly);
    }

    [Fact]
    public void SetPlacement_ConfiguresEmbeddedChartObjectProperties()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);
        var createResult = _commands.CreateFromRange(
            batch, "Sheet1", "A1:B6", ChartType.ColumnClustered, chartName: "ObjectPropertiesChart");

        var result = _commands.SetPlacement(
            batch,
            createResult.ChartName,
            placement: 2,
            printObject: false,
            locked: false,
            roundedCorners: true);

        Assert.True(result.Success, result.ErrorMessage);
        var properties = ReadChartObjectProperties(batch, createResult.ChartName);
        Assert.Equal(2, properties.Placement);
        Assert.False(properties.PrintObject);
        Assert.False(properties.Locked);
        Assert.True(properties.RoundedCorners);
    }

    [Fact]
    public void SetAreaFormat_AppliesChartAreaFillAndBorder()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);
        var createResult = _commands.CreateFromRange(
            batch, "Sheet1", "A1:B6", ChartType.ColumnClustered, chartName: "AreaFormatChart");

        var result = _commands.SetAreaFormat(
            batch,
            createResult.ChartName,
            ChartAreaTarget.Chart,
            fillColor: "#FF0000",
            fillTransparency: 0.25,
            lineColor: "#0000FF",
            lineWeight: 2.5);

        Assert.True(result.Success, result.ErrorMessage);
        var format = ReadChartAreaFormat(batch, createResult.ChartName);
        Assert.Equal(0x0000FF, format.FillColor);
        Assert.Equal(0.25f, format.FillTransparency, precision: 2);
        Assert.Equal(0xFF0000, format.LineColor);
        Assert.Equal(2.5f, format.LineWeight, precision: 2);
    }

    [Fact]
    public void SetSeriesFormat_AppliesSeriesFillAndLineMaterialFormatting()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.SharedTestFile);
        var createResult = _commands.CreateFromRange(
            batch, "Sheet1", "A1:B6", ChartType.ColumnClustered, chartName: "SeriesMaterialChart");

        var result = _commands.SetSeriesFormat(
            batch,
            createResult.ChartName,
            seriesIndex: 1,
            fillColor: "#00FF00",
            fillTransparency: 0.4,
            lineColor: "#FF00FF",
            lineWeight: 3);

        Assert.True(result.Success, result.ErrorMessage);
        var format = ReadSeriesFormat(batch, createResult.ChartName, 1);
        Assert.Equal(0x00FF00, format.FillColor);
        Assert.Equal(0.4f, format.FillTransparency, precision: 2);
        Assert.Equal(0xFF00FF, format.LineColor);
        Assert.Equal(3f, format.LineWeight, precision: 2);
    }

    private static ChartType ReadSeriesChartType(IExcelBatch batch, string chartName, int seriesIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.ChartObject? chartObject = null;
            Excel.Chart? chart = null;
            Excel.SeriesCollection? seriesCollection = null;
            Excel.Series? series = null;
            try
            {
                sheet = (Excel.Worksheet)ctx.Book.Worksheets["Sheet1"];
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
    }

    private static (int Placement, bool PrintObject, bool Locked, bool RoundedCorners) ReadChartObjectProperties(
        IExcelBatch batch,
        string chartName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.ChartObject? chartObject = null;
            try
            {
                sheet = (Excel.Worksheet)ctx.Book.Worksheets["Sheet1"];
                chartObject = (Excel.ChartObject)sheet.ChartObjects(chartName);
                return (
                    Convert.ToInt32((object)chartObject.Placement, CultureInfo.InvariantCulture),
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
    }

    private static (int FillColor, float FillTransparency, int LineColor, float LineWeight) ReadChartAreaFormat(
        IExcelBatch batch,
        string chartName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.ChartObject? chartObject = null;
            Excel.Chart? chart = null;
            Excel.ChartArea? chartArea = null;
            // Excel PIA exposes these Office drawing objects without a Microsoft.Office.Core reference.
            dynamic? chartFormat = null;
            dynamic? fill = null;
            dynamic? line = null;
            dynamic? fillColor = null;
            dynamic? lineColor = null;
            try
            {
                sheet = (Excel.Worksheet)ctx.Book.Worksheets["Sheet1"];
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
    }

    private static (int FillColor, float FillTransparency, int LineColor, float LineWeight) ReadSeriesFormat(
        IExcelBatch batch,
        string chartName,
        int seriesIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.ChartObject? chartObject = null;
            Excel.Chart? chart = null;
            Excel.SeriesCollection? seriesCollection = null;
            Excel.Series? series = null;
            // Excel PIA exposes these Office drawing objects without a Microsoft.Office.Core reference.
            dynamic? chartFormat = null;
            dynamic? fill = null;
            dynamic? line = null;
            dynamic? fillColor = null;
            dynamic? lineColor = null;
            try
            {
                sheet = (Excel.Worksheet)ctx.Book.Worksheets["Sheet1"];
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
}
