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

        // Get linked PivotTable name
        try
        {
            dynamic pivotLayout = chart.PivotLayout;
            dynamic pivotTable = pivotLayout.PivotTable;
            info.LinkedPivotTable = pivotTable.Name?.ToString() ?? string.Empty;
            ComUtilities.Release(ref pivotTable!);
            ComUtilities.Release(ref pivotLayout!);
        }
        catch
        {
            // No linked PivotTable
        }

        // Series count = number of value fields in PivotTable
        try
        {
            dynamic pivotLayout = chart.PivotLayout;
            dynamic pivotTable = pivotLayout.PivotTable;
            dynamic dataFields = pivotTable.DataFields;
            info.SeriesCount = Convert.ToInt32(dataFields.Count);
            ComUtilities.Release(ref dataFields!);
            ComUtilities.Release(ref pivotTable!);
            ComUtilities.Release(ref pivotLayout!);
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
            IsPivotChart = true,
            Left = Convert.ToDouble(shape.Left),
            Top = Convert.ToDouble(shape.Top),
            Width = Convert.ToDouble(shape.Width),
            Height = Convert.ToDouble(shape.Height)
        };

        // Get linked PivotTable name
        try
        {
            dynamic pivotLayout = chart.PivotLayout;
            dynamic pivotTable = pivotLayout.PivotTable;
            info.LinkedPivotTable = pivotTable.Name?.ToString() ?? string.Empty;
            ComUtilities.Release(ref pivotTable!);
            ComUtilities.Release(ref pivotLayout!);
        }
        catch
        {
            // No linked PivotTable
        }

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

        // PivotCharts don't expose series in the same way - data comes from PivotTable value fields
        // Series list remains empty for PivotCharts

        return info;
    }

    /// <inheritdoc />
    public void SetSourceRange(dynamic chart, string sourceRange)
    {
        // PivotCharts can't change source - throw helpful exception
        string? pivotTableName;
        dynamic? pivotLayout = null;
        dynamic? pivotTable = null;

        try
        {
            pivotLayout = chart.PivotLayout;
            pivotTable = pivotLayout.PivotTable;
            pivotTableName = pivotTable.Name?.ToString() ?? string.Empty;
        }
        catch
        {
            pivotTableName = "(unknown)";
        }
        finally
        {
            if (pivotTable != null) ComUtilities.Release(ref pivotTable!);
            if (pivotLayout != null) ComUtilities.Release(ref pivotLayout!);
        }

        throw new NotSupportedException(
            $"Cannot set source range for PivotChart. " +
            $"PivotCharts automatically sync with their PivotTable data source. " +
            $"To modify data, use excel_pivottable tool to update PivotTable '{pivotTableName}'.");
    }

    /// <inheritdoc />
    public SeriesInfo AddSeries(dynamic chart, string seriesName, string valuesRange, string? categoryRange)
    {
        // PivotCharts auto-sync with PivotTable fields - throw helpful exception
        string? pivotTableName;
        dynamic? pivotLayout = null;
        dynamic? pivotTable = null;

        try
        {
            pivotLayout = chart.PivotLayout;
            pivotTable = pivotLayout.PivotTable;
            pivotTableName = pivotTable.Name?.ToString() ?? string.Empty;
        }
        catch
        {
            pivotTableName = "(unknown)";
        }
        finally
        {
            if (pivotTable != null) ComUtilities.Release(ref pivotTable!);
            if (pivotLayout != null) ComUtilities.Release(ref pivotLayout!);
        }

        throw new NotSupportedException(
            $"Cannot add series directly to PivotChart. " +
            $"PivotCharts automatically sync with PivotTable '{pivotTableName}' fields. " +
            $"Use excel_pivottable(action: 'add-value-field', pivotTableName: '{pivotTableName}', fieldName: '<field>') " +
            $"to add data series.");
    }

    /// <inheritdoc />
    public void RemoveSeries(dynamic chart, int seriesIndex)
    {
        // PivotCharts auto-sync with PivotTable fields - throw helpful exception
        string? pivotTableName;
        dynamic? pivotLayout = null;
        dynamic? pivotTable = null;

        try
        {
            pivotLayout = chart.PivotLayout;
            pivotTable = pivotLayout.PivotTable;
            pivotTableName = pivotTable.Name?.ToString() ?? string.Empty;
        }
        catch
        {
            pivotTableName = "(unknown)";
        }
        finally
        {
            if (pivotTable != null) ComUtilities.Release(ref pivotTable!);
            if (pivotLayout != null) ComUtilities.Release(ref pivotLayout!);
        }

        throw new NotSupportedException(
            $"Cannot remove series directly from PivotChart. " +
            $"PivotCharts automatically sync with PivotTable '{pivotTableName}' fields. " +
            $"Use excel_pivottable(action: 'remove-field', pivotTableName: '{pivotTableName}', fieldName: '<field>') " +
            $"to remove data series.");
    }
}
