using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Strategy for Regular Charts (created from ranges/tables).
/// Handles Shapes.AddChart(), SeriesCollection operations, explicit data source management.
/// </summary>
public class RegularChartStrategy : IChartStrategy
{
    /// <inheritdoc />
    public bool CanHandle(Excel.Chart chart) =>
        !PivotChartStrategy.HasPivotLayout(() => chart.PivotLayout);

    /// <inheritdoc />
    public ChartInfo GetInfo(Excel.Chart chart, string chartName, string sheetName, Excel.Shape shape)
    {
        var info = new ChartInfo
        {
            Name = chartName,
            SheetName = sheetName,
            ChartType = (ChartType)Convert.ToInt32((object)chart.ChartType, System.Globalization.CultureInfo.InvariantCulture),
            IsPivotChart = false,
            Left = Convert.ToDouble(shape.Left),
            Top = Convert.ToDouble(shape.Top),
            Width = Convert.ToDouble(shape.Width),
            Height = Convert.ToDouble(shape.Height)
        };

        // Get anchor cells and placement mode
        Excel.Range? topLeftCell = null;
        Excel.Range? bottomRightCell = null;
        try
        {
            topLeftCell = shape.TopLeftCell;
            info.TopLeftCell = topLeftCell.Address?.ToString();
        }
        catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex))
        {
            // TopLeftCell not available - optional COM property
        }
        finally
        {
            ComUtilities.Release(ref topLeftCell!);
        }

        try
        {
            bottomRightCell = shape.BottomRightCell;
            info.BottomRightCell = bottomRightCell.Address?.ToString();
        }
        catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex))
        {
            // BottomRightCell not available - optional COM property
        }
        finally
        {
            ComUtilities.Release(ref bottomRightCell!);
        }

        try
        {
            info.Placement = (int)shape.Placement;
        }
        catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex))
        {
            // Placement not available - optional COM property
        }

        info.SeriesCount = ChartSeriesReader.Count(chart);

        return info;
    }

    /// <inheritdoc />
    public ChartInfoResult GetDetailedInfo(Excel.Chart chart, string chartName, string sheetName, Excel.Shape shape)
    {
        var info = new ChartInfoResult
        {
            Success = true,
            Name = chartName,
            SheetName = sheetName,
            ChartType = (ChartType)Convert.ToInt32((object)chart.ChartType, System.Globalization.CultureInfo.InvariantCulture),
            IsPivotChart = false,
            Left = Convert.ToDouble(shape.Left),
            Top = Convert.ToDouble(shape.Top),
            Width = Convert.ToDouble(shape.Width),
            Height = Convert.ToDouble(shape.Height)
        };

        // Get anchor cells and placement mode
        Excel.Range? topLeftCell = null;
        Excel.Range? bottomRightCell = null;
        try
        {
            topLeftCell = shape.TopLeftCell;
            info.TopLeftCell = topLeftCell.Address?.ToString();
        }
        catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex))
        {
            // TopLeftCell not available - optional COM property
        }
        finally
        {
            ComUtilities.Release(ref topLeftCell!);
        }

        try
        {
            bottomRightCell = shape.BottomRightCell;
            info.BottomRightCell = bottomRightCell.Address?.ToString();
        }
        catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex))
        {
            // BottomRightCell not available - optional COM property
        }
        finally
        {
            ComUtilities.Release(ref bottomRightCell!);
        }

        try
        {
            info.Placement = (int)shape.Placement;
        }
        catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex))
        {
            // Placement not available - optional COM property
        }

        // Get title
        if (chart.HasTitle)
        {
            Excel.ChartTitle? chartTitle = null;
            try
            {
                chartTitle = chart.ChartTitle;
                info.Title = chartTitle.Text?.ToString() ?? string.Empty;
            }
            finally
            {
                ComUtilities.Release(ref chartTitle);
            }
        }

        // Get legend
        info.HasLegend = chart.HasLegend;

        // Get source range
        Excel.ChartArea? chartArea = null;
        Excel.Chart? chartParent = null;
        Excel.SeriesCollection? sourceSeriesCollection = null;
        Excel.Series? sourceSeries = null;
        try
        {
            chartArea = chart.ChartArea;
            chartParent = (Excel.Chart)chartArea.Parent;
            sourceSeriesCollection = (Excel.SeriesCollection)chartParent.SeriesCollection();
            if (Convert.ToInt32(sourceSeriesCollection.Count) > 0)
            {
                sourceSeries = sourceSeriesCollection.Item(1);
                info.SourceRange = sourceSeries.Formula?.ToString() ?? string.Empty;
            }
        }
        finally
        {
            ComUtilities.Release(ref sourceSeries);
            ComUtilities.Release(ref sourceSeriesCollection);
            ComUtilities.Release(ref chartParent);
            ComUtilities.Release(ref chartArea);
        }

        info.Series = ChartSeriesReader.Read(chart);

        return info;
    }

    /// <inheritdoc />
    public void SetSourceRange(Excel.Chart chart, string sourceRange, CancellationToken ct)
    {
        Excel.ChartObject? chartObject = null;
        Excel.Worksheet? chartSheet = null;
        Excel.Workbook? book = null;
        Excel.Range? sourceRangeObj = null;
        try
        {
            chartObject = (Excel.ChartObject)chart.Parent;
            chartSheet = (Excel.Worksheet)chartObject.Parent;
            book = (Excel.Workbook)chartSheet.Parent;
            sourceRangeObj = ChartSourceRange.Resolve(book, chartSheet.Name, sourceRange, ct);
            ChartSourceRange.Apply(chart, sourceRangeObj, sourceRange);
        }
        finally
        {
            ComUtilities.Release(ref sourceRangeObj);
            ComUtilities.Release(ref book);
            ComUtilities.Release(ref chartSheet);
            ComUtilities.Release(ref chartObject);
        }
    }

    /// <inheritdoc />
    public SeriesInfo AddSeries(Excel.Chart chart, string seriesName, string valuesRange, string? categoryRange)
    {
        Excel.ChartObject? chartObject = null;
        Excel.Worksheet? chartSheet = null;
        Excel.Range? valuesSource = null;
        Excel.Range? categorySource = null;
        Excel.SeriesCollection? seriesCollection = null;
        Excel.Series? newSeries = null;

        try
        {
            chartObject = (Excel.ChartObject)chart.Parent;
            chartSheet = (Excel.Worksheet)chartObject.Parent;
            valuesSource = ResolveSeriesRange(chartSheet, valuesRange);
            if (!string.IsNullOrWhiteSpace(categoryRange))
            {
                categorySource = ResolveSeriesRange(chartSheet, categoryRange);
            }

            seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection();
            newSeries = seriesCollection.NewSeries();
            newSeries.Name = seriesName;
            newSeries.Values = valuesSource;

            if (categorySource != null)
            {
                newSeries.XValues = categorySource;
            }

            return new SeriesInfo
            {
                Name = seriesName,
                ValuesRange = valuesRange,
                CategoryRange = categoryRange
            };
        }
        finally
        {
            ComUtilities.Release(ref newSeries);
            ComUtilities.Release(ref seriesCollection);
            ComUtilities.Release(ref categorySource);
            ComUtilities.Release(ref valuesSource);
            ComUtilities.Release(ref chartSheet);
            ComUtilities.Release(ref chartObject);
        }
    }

    private static Excel.Range ResolveSeriesRange(Excel.Worksheet chartSheet, string reference)
    {
        var separator = reference.LastIndexOf('!');
        if (separator < 0)
        {
            return chartSheet.Range[reference];
        }

        var sheetName = reference[..separator].Trim();
        if (sheetName.Length >= 2 && sheetName[0] == '\'' && sheetName[^1] == '\'')
        {
            sheetName = sheetName[1..^1].Replace("''", "'", StringComparison.Ordinal);
        }

        Excel.Workbook? book = null;
        Excel.Sheets? sheets = null;
        Excel.Worksheet? sourceSheet = null;
        try
        {
            book = (Excel.Workbook)chartSheet.Parent;
            sheets = book.Worksheets;
            sourceSheet = (Excel.Worksheet)sheets[sheetName];
            return sourceSheet.Range[reference[(separator + 1)..]];
        }
        finally
        {
            ComUtilities.Release(ref sourceSheet);
            ComUtilities.Release(ref sheets);
            ComUtilities.Release(ref book);
        }
    }

    /// <inheritdoc />
    public void RemoveSeries(Excel.Chart chart, int seriesIndex)
    {
        Excel.SeriesCollection? seriesCollection = null;
        Excel.Series? series = null;

        try
        {
            seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection();
            series = seriesCollection.Item(seriesIndex);
            series.Delete();
        }
        finally
        {
            if (series != null)
            {
                ComUtilities.Release(ref series!);
            }
            if (seriesCollection != null)
            {
                ComUtilities.Release(ref seriesCollection!);
            }
        }
    }
}
