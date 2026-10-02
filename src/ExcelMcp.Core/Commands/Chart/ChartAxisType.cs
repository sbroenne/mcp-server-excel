namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Chart axis types for setting axis titles, scales, and gridlines.
/// </summary>
public enum ChartAxisType
{
    /// <summary>Legacy alias for Category on the primary axis group.</summary>
    Primary,

    /// <summary>Legacy alias for Value on the primary axis group, not a secondary-group selector.</summary>
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

