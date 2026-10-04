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
    public bool CanHandle(dynamic chart)
    {
        // Regular charts: chart.PivotLayout is null or doesn't exist
        dynamic? pivotLayout = null;
        try
        {
            pivotLayout = chart.PivotLayout;
            return pivotLayout == null;
        }
        catch (COMException)
        {
            return true; // No PivotLayout property = Regular chart
        }
        finally
        {
            ComUtilities.Release(ref pivotLayout);
        }
    }

    /// <inheritdoc />
    public ChartInfo GetInfo(dynamic chart, string chartName, string sheetName, dynamic shape)
    {
        var info = new ChartInfo
        {
            Name = chartName,
            SheetName = sheetName,
            ChartType = (ChartType)Convert.ToInt32(chart.ChartType),
            IsPivotChart = false,
            Left = Convert.ToDouble(shape.Left),
            Top = Convert.ToDouble(shape.Top),
            Width = Convert.ToDouble(shape.Width),
            Height = Convert.ToDouble(shape.Height)
        };

        // Get anchor cells and placement mode
        dynamic? topLeftCell = null;
        dynamic? bottomRightCell = null;
        try
        {
            topLeftCell = shape.TopLeftCell;
            info.TopLeftCell = topLeftCell.Address?.ToString();
        }
        catch (COMException)
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
        catch (COMException)
        {
            // BottomRightCell not available - optional COM property
        }
        finally
        {
            ComUtilities.Release(ref bottomRightCell!);
        }

        try
        {
            info.Placement = Convert.ToInt32(shape.Placement);
        }
        catch (COMException)
        {
            // Placement not available - optional COM property
        }

        // Count series
        dynamic? seriesCollection = null;
        try
        {
            seriesCollection = chart.SeriesCollection();
            info.SeriesCount = Convert.ToInt32(seriesCollection.Count);
        }
        finally
        {
            ComUtilities.Release(ref seriesCollection!);
        }

        return info;
    }

    /// <inheritdoc />
    public ChartInfoResult GetDetailedInfo(dynamic chart, string chartName, string sheetName, dynamic shape)
    {
        var info = new ChartInfoResult
        {
            Success = true,
            Name = chartName,
            SheetName = sheetName,
            ChartType = (ChartType)Convert.ToInt32(chart.ChartType),
            IsPivotChart = false,
            Left = Convert.ToDouble(shape.Left),
            Top = Convert.ToDouble(shape.Top),
            Width = Convert.ToDouble(shape.Width),
            Height = Convert.ToDouble(shape.Height)
        };

        // Get anchor cells and placement mode
        dynamic? topLeftCell = null;
        dynamic? bottomRightCell = null;
        try
        {
            topLeftCell = shape.TopLeftCell;
            info.TopLeftCell = topLeftCell.Address?.ToString();
        }
        catch (COMException)
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
        catch (COMException)
        {
            // BottomRightCell not available - optional COM property
        }
        finally
        {
            ComUtilities.Release(ref bottomRightCell!);
        }

        try
        {
            info.Placement = Convert.ToInt32(shape.Placement);
        }
        catch (COMException)
        {
            // Placement not available - optional COM property
        }

        // Get title
        try
        {
            if (chart.HasTitle)
            {
                info.Title = chart.ChartTitle.Text?.ToString() ?? string.Empty;
            }
        }
        catch (COMException)
        {
            // No title - optional COM property, safe to ignore
        }

        // Get legend
        try
        {
            info.HasLegend = chart.HasLegend;
        }
        catch (COMException)
        {
            info.HasLegend = false; // Safe fallback for optional COM property
        }

        // Get source range
        dynamic? chartArea = null;
        dynamic? chartParent = null;
        dynamic? sourceSeriesCollection = null;
        dynamic? sourceSeries = null;
        try
        {
            chartArea = chart.ChartArea;
            chartParent = chartArea.Parent;
            sourceSeriesCollection = chartParent.SeriesCollection();
            sourceSeries = sourceSeriesCollection.Item(1);
            info.SourceRange = sourceSeries.Formula?.ToString() ?? string.Empty;
        }
        catch (COMException)
        {
            // No source range or no series - optional COM property, safe to ignore
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
    public void SetSourceRange(dynamic chart, string sourceRange)
    {
        dynamic? app = null;
        dynamic? sourceRangeObj = null;
        try
        {
            app = chart.Application;
            sourceRangeObj = app.Range(sourceRange);
            chart.SetSourceData(sourceRangeObj);
        }
        finally
        {
            ComUtilities.Release(ref sourceRangeObj);
            ComUtilities.Release(ref app);
        }
    }

    /// <inheritdoc />
    public SeriesInfo AddSeries(dynamic chart, string seriesName, string valuesRange, string? categoryRange)
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
    public void RemoveSeries(dynamic chart, int seriesIndex)
    {
        dynamic? seriesCollection = null;
        dynamic? series = null;

        try
        {
            seriesCollection = chart.SeriesCollection();
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
