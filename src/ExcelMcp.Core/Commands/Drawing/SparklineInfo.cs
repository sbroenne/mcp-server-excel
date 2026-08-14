namespace Sbroenne.ExcelMcp.Core.Commands.Drawing;

/// <summary>Information about an Excel sparkline group.</summary>
public sealed class SparklineInfo
{
    /// <summary>Containing worksheet.</summary>
    public string SheetName { get; set; } = string.Empty;
    /// <summary>Cell range containing the sparklines.</summary>
    public string LocationRange { get; set; } = string.Empty;
    /// <summary>Source data range.</summary>
    public string SourceRange { get; set; } = string.Empty;
    /// <summary>Sparkline chart type.</summary>
    public DrawingSparklineType SparklineType { get; set; }
    /// <summary>Series color as #RRGGBB.</summary>
    public string? LineColor { get; set; }
    /// <summary>Whether line markers are displayed.</summary>
    public bool ShowMarkers { get; set; }
}
