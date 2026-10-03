using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceChartFormattingTests
{
    [Theory]
    [InlineData(ChartAxisType.Category, ChartAxisType.CategorySecondary, Excel.XlAxisType.xlCategory)]
    [InlineData(ChartAxisType.Value, ChartAxisType.ValueSecondary, Excel.XlAxisType.xlValue)]
    public void SecondaryAxisTitleAndNumberFormat_PreservePrimaryAxis(
        ChartAxisType primary, ChartAxisType secondary, Excel.XlAxisType axisType)
    {
        var batch = _fixture.BatchToken;
        var created = _chartCommands.CreateFromRange(batch, _sheetName, "A1:C6", ChartType.Line);
        Assert.True(created.Success, created.ErrorMessage);
        WithChart(chart =>
        {
            Excel.SeriesCollection? seriesCollection = null;
            Excel.Series? series = null;
            try
            {
                seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection();
                series = seriesCollection.Item(2);
                series.AxisGroup = Excel.XlAxisGroup.xlSecondary;
                chart.HasAxis[axisType, Excel.XlAxisGroup.xlSecondary] = true;
            }
            finally
            {
                ComUtilities.Release(ref series);
                ComUtilities.Release(ref seriesCollection);
            }
        });

        Assert.True(_chartCommands.SetAxisTitle(batch, created.ChartName, primary, "Primary").Success);
        Assert.True(_chartCommands.SetAxisNumberFormat(batch, created.ChartName, primary, "0.0").Success);
        Assert.True(_chartCommands.SetAxisTitle(batch, created.ChartName, secondary, "Secondary").Success);
        Assert.True(_chartCommands.SetAxisNumberFormat(batch, created.ChartName, secondary, "0.00").Success);
        Assert.Equal("0.0", _chartCommands.GetAxisNumberFormat(batch, created.ChartName, primary));
        Assert.Equal("0.00", _chartCommands.GetAxisNumberFormat(batch, created.ChartName, secondary));
        WithChart(chart =>
        {
            VerifyTitle(chart, Excel.XlAxisGroup.xlPrimary, "Primary");
            VerifyTitle(chart, Excel.XlAxisGroup.xlSecondary, "Secondary");
        });

        void VerifyTitle(Excel.Chart chart, Excel.XlAxisGroup group, string expected)
        {
            Excel.Axis? axis = null;
            Excel.AxisTitle? title = null;
            try
            {
                axis = (Excel.Axis)chart.Axes(axisType, group);
                Assert.True(axis.HasTitle);
                title = axis.AxisTitle;
                Assert.Equal(expected, title.Text);
            }
            finally
            {
                ComUtilities.Release(ref title);
                ComUtilities.Release(ref axis);
            }
        }

        void WithChart(Action<Excel.Chart> verify) =>
            _fixture.ExecuteRawVerification((ctx, _) =>
            {
                Excel.Sheets? sheets = null;
                Excel.Worksheet? sheet = null;
                Excel.ChartObjects? objects = null;
                Excel.ChartObject? chartObject = null;
                Excel.Chart? chart = null;
                try
                {
                    sheets = ctx.Book.Worksheets;
                    sheet = (Excel.Worksheet)sheets[_sheetName];
                    objects = (Excel.ChartObjects)sheet.ChartObjects();
                    chartObject = objects.Item(created.ChartName);
                    chart = chartObject.Chart;
                    verify(chart);
                }
                finally
                {
                    ComUtilities.Release(ref chart);
                    ComUtilities.Release(ref chartObject);
                    ComUtilities.Release(ref objects);
                    ComUtilities.Release(ref sheet);
                    ComUtilities.Release(ref sheets);
                }
            });
    }
}
