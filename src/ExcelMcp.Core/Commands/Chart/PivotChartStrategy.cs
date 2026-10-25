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
    public bool CanHandle(Excel.Chart chart) =>
        HasPivotLayout(() => chart.PivotLayout);

    internal delegate Excel.PivotLayout? PivotLayoutReader();

    internal static bool HasPivotLayout(PivotLayoutReader readPivotLayout)
    {
        Excel.PivotLayout? pivotLayout = null;
        try
        {
            pivotLayout = readPivotLayout();
            return pivotLayout != null;
        }
        catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex))
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
            ChartType = (ChartType)Convert.ToInt32((object)chart.ChartType, System.Globalization.CultureInfo.InvariantCulture),
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
            ChartType = (ChartType)Convert.ToInt32((object)chart.ChartType, System.Globalization.CultureInfo.InvariantCulture),
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
        info.HasLegend = chart.HasLegend;

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
