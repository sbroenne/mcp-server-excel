using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

public partial class ChartCommands
{
    /// <inheritdoc />
    public ChartSeriesSettingsResult GetSeriesSettings(IExcelBatch batch, string chartName, int seriesIndex)
    {
        ValidateSeriesIndex(seriesIndex);
        return WithNativeSeries(batch, chartName, seriesIndex, false,
            (chart, series, _, _, _) => ReadSeriesSettings(batch, chartName, seriesIndex, series));
    }

    /// <inheritdoc />
    public ChartSeriesSettingsResult SetSeriesAxisGroup(IExcelBatch batch, string chartName, int seriesIndex, ChartAxisGroup axisGroup)
    {
        ValidateSeriesIndex(seriesIndex);
        if (!Enum.IsDefined(axisGroup)) throw new ArgumentOutOfRangeException(nameof(axisGroup));
        return WithNativeSeries(batch, chartName, seriesIndex, true, (_, series, _, _, _) =>
        {
            series.AxisGroup = (Excel.XlAxisGroup)axisGroup;
            if ((int)series.AxisGroup != (int)axisGroup)
                throw new InvalidOperationException($"Excel did not apply {axisGroup} axes to this series. Its chart type may not support that assignment.");
            return ReadSeriesSettings(batch, chartName, seriesIndex, series);
        });
    }

    private static void ValidateSeriesIndex(int index)
    {
        if (index < 1) throw new ArgumentOutOfRangeException(nameof(index), "Series indices start at one.");
    }

    private T WithNativeSeries<T>(IExcelBatch batch, string chartName, int index, bool requireRegular,
        Func<Excel.Chart, Excel.Series, Excel.Workbook, string, CancellationToken, T> operation)
    {
        return batch.Execute((ctx, ct) =>
        {
            var found = FindChart(ctx.Book, chartName);
            Excel.Chart? chart = null;
            Excel.SeriesCollection? collection = null;
            Excel.Series? series = null;
            try
            {
                chart = (Excel.Chart?)found.Chart
                    ?? throw new ArgumentException($"Chart '{chartName}' not found.");
                found.Chart = null;
                if (requireRegular && _pivotStrategy.CanHandle(chart))
                    throw new InvalidOperationException("This per-series operation is not supported for PivotCharts. Use pivottable_field for their data fields and chartconfig.set-chart-type for the whole chart.");
                collection = (Excel.SeriesCollection)chart.SeriesCollection();
                if (index > collection.Count)
                    throw new ArgumentOutOfRangeException(nameof(index), $"Chart has {collection.Count} series.");
                series = collection.Item(index);
                ct.ThrowIfCancellationRequested();
                return operation(chart, series, ctx.Book, found.SheetName, ct);
            }
            finally
            {
                ComUtilities.Release(ref series);
                ComUtilities.Release(ref collection);
                ComUtilities.Release(ref chart);
                if (found.Shape != null) ComUtilities.Release(ref found.Shape!);
                if (found.Chart != null) ComUtilities.Release(ref found.Chart!);
            }
        });
    }

    private static ChartSeriesSettingsResult ReadSeriesSettings(IExcelBatch batch, string name, int index, Excel.Series series)
    {
        Excel.Points? points = null;
        try
        {
            points = (Excel.Points)series.Points();
            return new ChartSeriesSettingsResult
            {
                Success = true,
                FilePath = batch.WorkbookPath,
                ChartName = name,
                SeriesIndex = index,
                Name = series.Name,
                Formula = series.Formula,
                ChartType = (ChartType)series.ChartType,
                AxisGroup = (ChartAxisGroup)series.AxisGroup,
                PointCount = points.Count,
                HasErrorBars = series.HasErrorBars
            };
        }
        finally
        {
            ComUtilities.Release(ref points);
        }
    }
}
