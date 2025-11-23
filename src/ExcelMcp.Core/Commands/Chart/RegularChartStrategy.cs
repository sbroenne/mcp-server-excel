using Sbroenne.ExcelMcp.ComInterop;

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
        try
        {
            var pivotLayout = chart.PivotLayout;
            return pivotLayout == null;
        }
        catch
        {
            return true; // No PivotLayout property = Regular chart
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

        // Count series
        try
        {
            dynamic seriesCollection = chart.SeriesCollection();
            info.SeriesCount = Convert.ToInt32(seriesCollection.Count);
            ComUtilities.Release(ref seriesCollection!);
        }
        catch
        {
            info.SeriesCount = 0;
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

        // Get title
        try
        {
            if (chart.HasTitle)
            {
                info.Title = chart.ChartTitle.Text?.ToString() ?? string.Empty;
            }
        }
        catch
        {
            // No title
        }

        // Get legend
        try
        {
            info.HasLegend = chart.HasLegend;
        }
        catch
        {
            info.HasLegend = false;
        }

        // Get source range
        try
        {
            dynamic sourceData = chart.ChartArea.Parent.SeriesCollection(1).Formula;
            info.SourceRange = sourceData?.ToString() ?? string.Empty;
        }
        catch
        {
            // No source range or no series
        }

        // Get series
        try
        {
            dynamic seriesCollection = chart.SeriesCollection();
            int seriesCount = Convert.ToInt32(seriesCollection.Count);

            for (int i = 1; i <= seriesCount; i++)
            {
                dynamic? series = null;
                try
                {
                    series = seriesCollection.Item(i);
                    var seriesInfo = new SeriesInfo
                    {
                        Name = series.Name?.ToString() ?? string.Empty,
                        ValuesRange = series.Values?.ToString() ?? string.Empty,
                        CategoryRange = series.XValues?.ToString() ?? string.Empty
                    };
                    info.Series.Add(seriesInfo);
                }
                finally
                {
                    if (series != null)
                    {
                        ComUtilities.Release(ref series!);
                    }
                }
            }

            ComUtilities.Release(ref seriesCollection!);
        }
        catch
        {
            // No series or error reading
        }

        return info;
    }

    /// <inheritdoc />
    public void SetSourceRange(dynamic chart, string sourceRange)
    {
        dynamic? sourceRangeObj = null;
        try
        {
            // Get workbook from chart
            dynamic workbook = chart.Parent.Parent.Parent;

            // Get the range object from the address string
            sourceRangeObj = workbook.Application.Range(sourceRange);
            chart.SetSourceData(sourceRangeObj);
        }
        finally
        {
            if (sourceRangeObj != null)
            {
                ComUtilities.Release(ref sourceRangeObj!);
            }
        }
    }

    /// <inheritdoc />
    public SeriesInfo AddSeries(dynamic chart, string seriesName, string valuesRange, string? categoryRange)
    {
        dynamic? seriesCollection = null;
        dynamic? newSeries = null;

        try
        {
            seriesCollection = chart.SeriesCollection();
            newSeries = seriesCollection.NewSeries();
            newSeries.Name = seriesName;
            newSeries.Values = valuesRange;

            if (!string.IsNullOrWhiteSpace(categoryRange))
            {
                newSeries.XValues = categoryRange;
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
            if (newSeries != null)
            {
                ComUtilities.Release(ref newSeries!);
            }
            if (seriesCollection != null)
            {
                ComUtilities.Release(ref seriesCollection!);
            }
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
