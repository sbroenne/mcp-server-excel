namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Types of trendlines for chart data series.
/// Maps to Excel's XlTrendlineType enum values.
/// </summary>
public enum TrendlineType
{
    /// <summary>Linear regression (y = mx + b)</summary>
    Linear = -4132,

    /// <summary>Exponential (y = ce^bx)</summary>
    Exponential = 5,

    /// <summary>Logarithmic (y = c ln x + b)</summary>
    Logarithmic = -4133,

    /// <summary>Polynomial - requires Order parameter (2-6)</summary>
    Polynomial = 3,

    /// <summary>Power (y = cx^b)</summary>
    Power = 4,

    /// <summary>Moving average - requires Period parameter</summary>
    MovingAverage = 6
}


