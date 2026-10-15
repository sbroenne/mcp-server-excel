using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Strategy for PivotCharts (created from PivotTables).
/// Handles PivotCache.CreatePivotChart(), automatic sync with PivotTable, helpful errors for series operations.
/// </summary>
public class PivotChartStrategy : IChartStrategy
{
    /// <inheritdoc />
    public bool CanHandle(Excel.Chart chart)
    {
        // PivotCharts: chart.PivotLayout exists
        Excel.PivotLayout? pivotLayout = null;
        try
        {
            pivotLayout = chart.PivotLayout;
            return pivotLayout != null;
        }
        catch (COMException)
        {
            return false;
        }
        finally
        {
            ComUtilities.Release(ref pivotLayout);
        }
    }

    /// <inheritdoc />
    public ChartInfo GetInfo(Excel.Chart chart, string chartName, string sheetName, Excel.Shape shape)
    {
        var info = new ChartInfo
        {
            Name = chartName,
            SheetName = sheetName,
            ChartType = (ChartType)(int)chart.ChartType,
            IsPivotChart = true,
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
            info.Placement = Convert.ToInt32(shape.Placement, System.Globalization.CultureInfo.InvariantCulture);
        }
        catch (COMException)
        {
            // Placement not available - optional COM property
        }

        // Get linked PivotTable name
        Excel.PivotLayout? pivotLayout = null;
        Excel.PivotTable? pivotTable = null;
        try
        {
            pivotLayout = chart.PivotLayout;
            pivotTable = pivotLayout.PivotTable;
            info.LinkedPivotTable = pivotTable.Name?.ToString() ?? string.Empty;
        }
        finally
        {
            ComUtilities.Release(ref pivotTable!);
            ComUtilities.Release(ref pivotLayout!);
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
            ChartType = (ChartType)(int)chart.ChartType,
            IsPivotChart = true,
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
            info.Placement = Convert.ToInt32(shape.Placement, System.Globalization.CultureInfo.InvariantCulture);
        }
        catch (COMException)
        {
            // Placement not available - optional COM property
        }

        // Get linked PivotTable name
        Excel.PivotLayout? pivotLayout = null;
        Excel.PivotTable? pivotTable = null;
        try
        {
            pivotLayout = chart.PivotLayout;
            pivotTable = pivotLayout.PivotTable;
            info.LinkedPivotTable = pivotTable.Name?.ToString() ?? string.Empty;
        }
        finally
        {
            ComUtilities.Release(ref pivotTable!);
            ComUtilities.Release(ref pivotLayout!);
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
                ComUtilities.Release(ref chartTitle!);
            }
        }

        // Get legend
        try
        {
            info.HasLegend = chart.HasLegend;
        }
        catch (COMException)
        {
            info.HasLegend = false;
        }

        info.Series = ChartSeriesReader.Read(chart);

        return info;
    }

    /// <inheritdoc />
    public void SetSourceRange(Excel.Chart chart, string sourceRange, CancellationToken ct)
    {
        throw new NotSupportedException(
            "Cannot set source range for PivotChart. " +
            "PivotCharts automatically sync with their PivotTable data source. " +
            "Use pivottable tool to update the linked PivotTable.");
    }

    /// <inheritdoc />
    public SeriesInfo AddSeries(Excel.Chart chart, string seriesName, string valuesRange, string? categoryRange)
    {
        throw new NotSupportedException(
            "Cannot add series directly to PivotChart. " +
            "PivotCharts automatically sync with PivotTable fields. " +
            "Use pivottable_field (CLI: pivottablefield) with 'add-value-field' to add data series.");
    }

    /// <inheritdoc />
    public void RemoveSeries(Excel.Chart chart, int seriesIndex)
    {
        throw new NotSupportedException(
            "Cannot remove series directly from PivotChart. " +
            "PivotCharts automatically sync with PivotTable fields. " +
            "Use pivottable_field (CLI: pivottablefield) with 'remove-field' to remove data series.");
    }
}
