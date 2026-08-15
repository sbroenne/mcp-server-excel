using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Whether Excel treats source rows or columns as chart series.
/// </summary>
public enum ChartPlotBy
{
    /// <summary>Each source row is a series.</summary>
    Rows = 1,

    /// <summary>Each source column is a series.</summary>
    Columns = 2
}

/// <summary>
/// How a chart renders blank source cells.
/// </summary>
public enum ChartDisplayBlanksAs
{
    /// <summary>Leave a gap for blank cells.</summary>
    Gaps = 1,

    /// <summary>Plot blank cells as zero.</summary>
    Zero = 2,

    /// <summary>Interpolate between neighboring points.</summary>
    Interpolated = 3
}

/// <summary>
/// Chart region targeted by area formatting.
/// </summary>
public enum ChartAreaTarget
{
    /// <summary>Entire chart canvas, including title and legend.</summary>
    Chart,

    /// <summary>Plot rectangle containing the data series.</summary>
    Plot
}

/// <summary>
/// Chart source-orientation and blank/hidden-cell plotting behavior.
/// </summary>
public class ChartPlotOptionsResult : ResultBase
{
    /// <summary>Name of the chart.</summary>
    public string ChartName { get; set; } = string.Empty;

    /// <summary>Whether source rows or columns define series.</summary>
    public ChartPlotBy PlotBy { get; set; }

    /// <summary>How blank source cells are rendered.</summary>
    public ChartDisplayBlanksAs DisplayBlanksAs { get; set; }

    /// <summary>True when hidden rows and columns are omitted.</summary>
    public bool PlotVisibleOnly { get; set; }
}
