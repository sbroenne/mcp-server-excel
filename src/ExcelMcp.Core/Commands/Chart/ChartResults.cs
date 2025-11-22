using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Result containing list of charts in workbook.
/// </summary>
public class ChartListResult : OperationResult
{
    /// <summary>
    /// List of charts (Regular and PivotCharts).
    /// </summary>
    public List<ChartInfo> Charts { get; set; } = new();
}

/// <summary>
/// Information about a chart.
/// </summary>
public class ChartInfo
{
    /// <summary>Chart or shape name</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Worksheet containing the chart</summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>Chart type (column, line, pie, etc.)</summary>
    public ChartType ChartType { get; set; }

    /// <summary>True if this is a PivotChart, false if Regular Chart</summary>
    public bool IsPivotChart { get; set; }

    /// <summary>Name of linked PivotTable (PivotCharts only)</summary>
    public string? LinkedPivotTable { get; set; }

    /// <summary>Left position in points</summary>
    public double Left { get; set; }

    /// <summary>Top position in points</summary>
    public double Top { get; set; }

    /// <summary>Width in points</summary>
    public double Width { get; set; }

    /// <summary>Height in points</summary>
    public double Height { get; set; }

    /// <summary>Number of data series</summary>
    public int SeriesCount { get; set; }
}

/// <summary>
/// Result containing complete chart configuration.
/// </summary>
public class ChartInfoResult : OperationResult
{
    /// <summary>Chart or shape name</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Worksheet containing the chart</summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>Chart type (column, line, pie, etc.)</summary>
    public ChartType ChartType { get; set; }

    /// <summary>True if this is a PivotChart, false if Regular Chart</summary>
    public bool IsPivotChart { get; set; }

    /// <summary>Name of linked PivotTable (PivotCharts only)</summary>
    public string? LinkedPivotTable { get; set; }

    /// <summary>Source range for Regular Charts (e.g., "Sheet1!A1:D10")</summary>
    public string? SourceRange { get; set; }

    /// <summary>Left position in points</summary>
    public double Left { get; set; }

    /// <summary>Top position in points</summary>
    public double Top { get; set; }

    /// <summary>Width in points</summary>
    public double Width { get; set; }

    /// <summary>Height in points</summary>
    public double Height { get; set; }

    /// <summary>Chart title text</summary>
    public string? Title { get; set; }

    /// <summary>True if legend is visible</summary>
    public bool HasLegend { get; set; }

    /// <summary>Data series (Regular Charts only)</summary>
    public List<SeriesInfo> Series { get; set; } = new();
}

/// <summary>
/// Information about a chart data series.
/// </summary>
public class SeriesInfo
{
    /// <summary>Series name</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Range containing Y values</summary>
    public string ValuesRange { get; set; } = string.Empty;

    /// <summary>Range containing X values/categories (optional)</summary>
    public string? CategoryRange { get; set; }
}

/// <summary>
/// Result from chart creation operations.
/// </summary>
public class ChartCreateResult : OperationResult
{
    /// <summary>Name of the created chart</summary>
    public string ChartName { get; set; } = string.Empty;

    /// <summary>Worksheet containing the chart</summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>Chart type</summary>
    public ChartType ChartType { get; set; }

    /// <summary>True if this is a PivotChart, false if Regular Chart</summary>
    public bool IsPivotChart { get; set; }

    /// <summary>Name of linked PivotTable (PivotCharts only)</summary>
    public string? LinkedPivotTable { get; set; }

    /// <summary>Left position in points</summary>
    public double Left { get; set; }

    /// <summary>Top position in points</summary>
    public double Top { get; set; }

    /// <summary>Width in points</summary>
    public double Width { get; set; }

    /// <summary>Height in points</summary>
    public double Height { get; set; }
}

/// <summary>
/// Result from series operations.
/// </summary>
public class ChartSeriesResult : OperationResult
{
    /// <summary>Series name</summary>
    public string SeriesName { get; set; } = string.Empty;

    /// <summary>Range containing Y values</summary>
    public string ValuesRange { get; set; } = string.Empty;

    /// <summary>Range containing X values/categories (optional)</summary>
    public string? CategoryRange { get; set; }

    /// <summary>1-based series index</summary>
    public int SeriesIndex { get; set; }
}
