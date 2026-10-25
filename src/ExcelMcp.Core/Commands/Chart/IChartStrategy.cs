using Excel = Microsoft.Office.Interop.Excel;
namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Strategy interface for handling differences between Regular Charts and PivotCharts.
/// Abstracts COM API differences while providing unified API surface.
/// </summary>
public interface IChartStrategy
{
    /// <summary>
    /// Determines if this strategy can handle the given chart.
    /// </summary>
    /// <param name="chart">Excel Chart COM object</param>
    /// <returns>True if this strategy handles this chart type</returns>
    bool CanHandle(Excel.Chart chart);

    /// <summary>
    /// Gets chart information (type, position, series, etc.).
    /// </summary>
    /// <param name="chart">Excel Chart COM object</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="shape">Shape containing the chart</param>
    /// <returns>Chart information</returns>
    ChartInfo GetInfo(Excel.Chart chart, string chartName, string sheetName, Excel.Shape shape);

    /// <summary>
    /// Sets the data source range.
    /// Regular Charts: Updates source range.
    /// PivotCharts: Throws exception guiding to pivottable.
    /// </summary>
    /// <param name="chart">Excel Chart COM object</param>
    /// <param name="sourceRange">New source range</param>
    void SetSourceRange(Excel.Chart chart, string sourceRange, CancellationToken ct);

    /// <summary>
    /// Adds a data series.
    /// Regular Charts: Adds to SeriesCollection.
    /// PivotCharts: Throws exception guiding to pivottable.
    /// </summary>
    /// <param name="chart">Excel Chart COM object</param>
    /// <param name="seriesName">Name for the series</param>
    /// <param name="valuesRange">Range containing Y values</param>
    /// <param name="categoryRange">Optional range for X values/categories</param>
    /// <returns>Series information</returns>
    SeriesInfo AddSeries(Excel.Chart chart, string seriesName, string valuesRange, string? categoryRange);

    /// <summary>
    /// Removes a data series.
    /// Regular Charts: Removes from SeriesCollection.
    /// PivotCharts: Throws exception guiding to pivottable.
    /// </summary>
    /// <param name="chart">Excel Chart COM object</param>
    /// <param name="seriesIndex">1-based series index</param>
    void RemoveSeries(Excel.Chart chart, int seriesIndex);

    /// <summary>
    /// Gets detailed chart information including series.
    /// </summary>
    /// <param name="chart">Excel Chart COM object</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="shape">Shape containing the chart</param>
    /// <returns>Detailed chart information</returns>
    ChartInfoResult GetDetailedInfo(Excel.Chart chart, string chartName, string sheetName, Excel.Shape shape);
}
