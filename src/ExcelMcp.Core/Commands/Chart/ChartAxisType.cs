namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Chart axis types for setting axis titles, scales, and gridlines.
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
    Value,

    /// <summary>Secondary category axis (X-axis on secondary axis group)</summary>
    CategorySecondary,

    /// <summary>Secondary value axis (Y-axis on secondary axis group)</summary>
    ValueSecondary
}


