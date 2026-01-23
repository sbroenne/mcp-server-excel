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

/// <summary>
/// Result containing axis scale information.
/// </summary>
public class AxisScaleResult : OperationResult
{
    /// <summary>Chart name</summary>
    public string ChartName { get; set; } = string.Empty;

    /// <summary>Axis type (Value, Category, ValueSecondary, CategorySecondary)</summary>
    public string AxisType { get; set; } = string.Empty;

    /// <summary>Minimum scale value (null if auto)</summary>
    public double? MinimumScale { get; set; }

    /// <summary>Maximum scale value (null if auto)</summary>
    public double? MaximumScale { get; set; }

    /// <summary>True if minimum scale is automatic</summary>
    public bool MinimumScaleIsAuto { get; set; }

    /// <summary>True if maximum scale is automatic</summary>
    public bool MaximumScaleIsAuto { get; set; }

    /// <summary>Major unit (distance between major gridlines/tick marks)</summary>
    public double? MajorUnit { get; set; }

    /// <summary>Minor unit (distance between minor gridlines/tick marks)</summary>
    public double? MinorUnit { get; set; }

    /// <summary>True if major unit is automatic</summary>
    public bool MajorUnitIsAuto { get; set; }

    /// <summary>True if minor unit is automatic</summary>
    public bool MinorUnitIsAuto { get; set; }
}

/// <summary>
/// Information about chart data labels.
/// </summary>
public class DataLabelsInfo
{
    /// <summary>Show the actual value</summary>
    public bool ShowValue { get; set; }

    /// <summary>Show percentage (pie/doughnut charts)</summary>
    public bool ShowPercentage { get; set; }

    /// <summary>Show series name</summary>
    public bool ShowSeriesName { get; set; }

    /// <summary>Show category name</summary>
    public bool ShowCategoryName { get; set; }

    /// <summary>Show bubble size (bubble charts)</summary>
    public bool ShowBubbleSize { get; set; }

    /// <summary>Separator between label parts (e.g., ", " or newline)</summary>
    public string? Separator { get; set; }

    /// <summary>Position of data labels</summary>
    public string? Position { get; set; }
}

/// <summary>
/// Information about chart gridlines.
/// </summary>
public class GridlinesInfo
{
    /// <summary>True if major gridlines are visible on primary value axis</summary>
    public bool HasValueMajorGridlines { get; set; }

    /// <summary>True if minor gridlines are visible on primary value axis</summary>
    public bool HasValueMinorGridlines { get; set; }

    /// <summary>True if major gridlines are visible on category axis</summary>
    public bool HasCategoryMajorGridlines { get; set; }

    /// <summary>True if minor gridlines are visible on category axis</summary>
    public bool HasCategoryMinorGridlines { get; set; }
}

/// <summary>
/// Result containing gridlines information.
/// </summary>
public class GridlinesResult : OperationResult
{
    /// <summary>Chart name</summary>
    public string ChartName { get; set; } = string.Empty;

    /// <summary>Gridlines configuration</summary>
    public GridlinesInfo Gridlines { get; set; } = new();
}

/// <summary>
/// Information about series marker formatting.
/// </summary>
public class SeriesFormatInfo
{
    /// <summary>1-based series index</summary>
    public int SeriesIndex { get; set; }

    /// <summary>Series name</summary>
    public string SeriesName { get; set; } = string.Empty;

    /// <summary>Marker style (none, square, diamond, triangle, x, star, circle, plus, etc.)</summary>
    public string? MarkerStyle { get; set; }

    /// <summary>Marker size (2-72 points)</summary>
    public int? MarkerSize { get; set; }

    /// <summary>Marker background color (#RRGGBB hex)</summary>
    public string? MarkerBackgroundColor { get; set; }

    /// <summary>Marker foreground/border color (#RRGGBB hex)</summary>
    public string? MarkerForegroundColor { get; set; }

    /// <summary>True to invert colors for negative values</summary>
    public bool? InvertIfNegative { get; set; }
}

/// <summary>
/// Information about a trendline on a chart series.
/// </summary>
public class TrendlineInfo
{
    /// <summary>1-based trendline index within the series</summary>
    public int Index { get; set; }

    /// <summary>Trendline type (Linear, Exponential, etc.)</summary>
    public TrendlineType Type { get; set; }

    /// <summary>Custom name for the trendline</summary>
    public string? Name { get; set; }

    /// <summary>Polynomial order (2-6) when type is Polynomial</summary>
    public int? Order { get; set; }

    /// <summary>Moving average period when type is MovingAverage</summary>
    public int? Period { get; set; }

    /// <summary>Number of periods to forecast forward</summary>
    public double? Forward { get; set; }

    /// <summary>Number of periods to forecast backward</summary>
    public double? Backward { get; set; }

    /// <summary>Y-intercept value (null = calculated)</summary>
    public double? Intercept { get; set; }

    /// <summary>True if equation is displayed on chart</summary>
    public bool DisplayEquation { get; set; }

    /// <summary>True if R-squared value is displayed on chart</summary>
    public bool DisplayRSquared { get; set; }
}

/// <summary>
/// Result containing list of trendlines for a series.
/// </summary>
public class TrendlineListResult : OperationResult
{
    /// <summary>Chart name</summary>
    public string ChartName { get; set; } = string.Empty;

    /// <summary>1-based series index</summary>
    public int SeriesIndex { get; set; }

    /// <summary>Series name</summary>
    public string SeriesName { get; set; } = string.Empty;

    /// <summary>List of trendlines on the series</summary>
    public List<TrendlineInfo> Trendlines { get; set; } = new();
}

/// <summary>
/// Result from adding a trendline.
/// </summary>
public class TrendlineResult : OperationResult
{
    /// <summary>Chart name</summary>
    public string ChartName { get; set; } = string.Empty;

    /// <summary>1-based series index</summary>
    public int SeriesIndex { get; set; }

    /// <summary>1-based trendline index within the series</summary>
    public int TrendlineIndex { get; set; }

    /// <summary>Trendline type</summary>
    public TrendlineType Type { get; set; }

    /// <summary>Custom name for the trendline</summary>
    public string? Name { get; set; }
}

