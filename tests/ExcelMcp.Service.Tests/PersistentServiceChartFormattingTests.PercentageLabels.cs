using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceChartFormattingTests
{
    [Theory]
    [InlineData(ChartType.Pie)]
    [InlineData(ChartType.Doughnut)]
    public void SetDataLabels_PercentageOnSupportedChart_PreservesActualExcelFlags(ChartType type)
    {
        var batch = _fixture.BatchToken;
        var created = RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4", type));
        RequireSuccess(_chartCommands.SetDataLabels(
            batch, created.ChartName, showValue: false, showPercentage: true));
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.ChartObjects? charts = null;
            Excel.ChartObject? chartObject = null;
            Excel.Chart? chart = null;
            Excel.SeriesCollection? seriesCollection = null;
            Excel.Series? series = null;
            Excel.DataLabels? labels = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[_sheetName];
                charts = (Excel.ChartObjects)sheet.ChartObjects();
                chartObject = charts.Item(created.ChartName);
                chart = chartObject.Chart;
                seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection();
                series = seriesCollection.Item(1);
                Assert.True(series.HasDataLabels);
                labels = (Excel.DataLabels)series.DataLabels();
                Assert.True(labels.ShowPercentage);
                Assert.False(labels.ShowValue);
            }
            finally
            {
                ComUtilities.Release(ref labels);
                ComUtilities.Release(ref series);
                ComUtilities.Release(ref seriesCollection);
                ComUtilities.Release(ref chart);
                ComUtilities.Release(ref chartObject);
                ComUtilities.Release(ref charts);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }
}
