namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Chart axis types for setting axis titles.
/// </summary>
public enum ChartAxisType
{
    /// <summary>Primary horizontal axis (category axis for most charts)</summary>
    Primary,

    /// <summary>Secondary horizontal axis</summary>
    Secondary,

    /// <summary>Category axis (X-axis)</summary>
    Category,

    /// <summary>Value axis (Y-axis)</summary>
    Value
}
