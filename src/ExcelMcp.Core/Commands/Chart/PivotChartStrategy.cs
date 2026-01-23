using Sbroenne.ExcelMcp.ComInterop;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Strategy for PivotCharts (created from PivotTables).
/// Handles PivotCache.CreatePivotChart(), automatic sync with PivotTable, helpful errors for series operations.
/// </summary>
public class PivotChartStrategy : IChartStrategy
{
    /// <inheritdoc />
    public bool CanHandle(dynamic chart)
    {
        // PivotCharts: chart.PivotLayout exists
        try
        {
            var pivotLayout = chart.PivotLayout;
            return pivotLayout != null;
        }
        catch
        {
            return false;
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
            IsPivotChart = true,
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
        catch
        {
            // TopLeftCell not available - optional property
        }
        finally
        {
            if (topLeftCell != null) ComUtilities.Release(ref topLeftCell!);
        }

        try
        {
            bottomRightCell = shape.BottomRightCell;
            info.BottomRightCell = bottomRightCell.Address?.ToString();
        }
        catch
        {
            // BottomRightCell not available - optional property
        }
        finally
        {
            if (bottomRightCell != null) ComUtilities.Release(ref bottomRightCell!);
        }

        try
        {
            info.Placement = Convert.ToInt32(shape.Placement);
        }
        catch
        {
            // Placement not available - optional property
        }

        // Get linked PivotTable name
        dynamic? pivotLayout = null;
        dynamic? pivotTable = null;
        try
        {
            pivotLayout = chart.PivotLayout;
            pivotTable = pivotLayout.PivotTable;
            info.LinkedPivotTable = pivotTable.Name?.ToString() ?? string.Empty;
        }
        finally
        {
            if (pivotTable != null) ComUtilities.Release(ref pivotTable!);
            if (pivotLayout != null) ComUtilities.Release(ref pivotLayout!);
        }

        // Series count = number of value fields in PivotTable
        dynamic? pivotLayout2 = null;
        dynamic? pivotTable2 = null;
        dynamic? dataFields = null;
        try
        {
            pivotLayout2 = chart.PivotLayout;
            pivotTable2 = pivotLayout2.PivotTable;
            dataFields = pivotTable2.DataFields;
            info.SeriesCount = Convert.ToInt32(dataFields.Count);
        }
        finally
        {
            if (dataFields != null) ComUtilities.Release(ref dataFields!);
            if (pivotTable2 != null) ComUtilities.Release(ref pivotTable2!);
            if (pivotLayout2 != null) ComUtilities.Release(ref pivotLayout2!);
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
            IsPivotChart = true,
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
        catch
        {
            // TopLeftCell not available - optional property
        }
        finally
        {
            if (topLeftCell != null) ComUtilities.Release(ref topLeftCell!);
        }

        try
        {
            bottomRightCell = shape.BottomRightCell;
            info.BottomRightCell = bottomRightCell.Address?.ToString();
        }
        catch
        {
            // BottomRightCell not available - optional property
        }
        finally
        {
            if (bottomRightCell != null) ComUtilities.Release(ref bottomRightCell!);
        }

        try
        {
            info.Placement = Convert.ToInt32(shape.Placement);
        }
        catch
        {
            // Placement not available - optional property
        }

        // Get linked PivotTable name
        dynamic? pivotLayout = null;
        dynamic? pivotTable = null;
        try
        {
            pivotLayout = chart.PivotLayout;
            pivotTable = pivotLayout.PivotTable;
            info.LinkedPivotTable = pivotTable.Name?.ToString() ?? string.Empty;
        }
        finally
        {
            if (pivotTable != null) ComUtilities.Release(ref pivotTable!);
            if (pivotLayout != null) ComUtilities.Release(ref pivotLayout!);
        }

        // Get title
        if (chart.HasTitle)
        {
            info.Title = chart.ChartTitle.Text?.ToString() ?? string.Empty;
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

        // PivotCharts don't expose series in the same way - data comes from PivotTable value fields
        // Series list remains empty for PivotCharts

        return info;
    }

    /// <inheritdoc />
    public void SetSourceRange(dynamic chart, string sourceRange)
    {
        throw new NotSupportedException(
            "Cannot set source range for PivotChart. " +
            "PivotCharts automatically sync with their PivotTable data source. " +
            "Use excel_pivottable tool to update the linked PivotTable.");
    }

    /// <inheritdoc />
    public SeriesInfo AddSeries(dynamic chart, string seriesName, string valuesRange, string? categoryRange)
    {
        throw new NotSupportedException(
            "Cannot add series directly to PivotChart. " +
            "PivotCharts automatically sync with PivotTable fields. " +
            "Use excel_pivottable tool with 'add-value-field' action to add data series.");
    }

    /// <inheritdoc />
    public void RemoveSeries(dynamic chart, int seriesIndex)
    {
        throw new NotSupportedException(
            "Cannot remove series directly from PivotChart. " +
            "PivotCharts automatically sync with PivotTable fields. " +
            "Use excel_pivottable tool with 'remove-field' action to remove data series.");
    }
}
