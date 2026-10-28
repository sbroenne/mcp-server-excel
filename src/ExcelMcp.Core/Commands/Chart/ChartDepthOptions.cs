using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>Native series axis assignment.</summary>
public enum ChartAxisGroup
{
    /// <summary>Primary axes.</summary>
    Primary = 1,
    /// <summary>Secondary axes.</summary>
    Secondary = 2
}

/// <summary>Native selected-series settings.</summary>
public sealed class ChartSeriesSettingsResult : OperationResult
{
    /// <summary>Selected chart.</summary>
    public string ChartName { get; set; } = string.Empty;
    /// <summary>One-based selected series index.</summary>
    public int SeriesIndex { get; set; }
    /// <summary>Actual native series name.</summary>
    public string Name { get; set; } = string.Empty;
    /// <summary>Native series chart type.</summary>
    public ChartType ChartType { get; set; }
    /// <summary>Native primary/secondary axis assignment.</summary>
    public ChartAxisGroup AxisGroup { get; set; }
    /// <summary>Native SERIES formula, including source references.</summary>
    public string Formula { get; set; } = string.Empty;
    /// <summary>Native point count.</summary>
    public int PointCount { get; set; }
    /// <summary>Whether Excel reports error bars.</summary>
    public bool HasErrorBars { get; set; }
}

/// <summary>Native error-bar calculation.</summary>
public enum ChartErrorBarKind
{
    /// <summary>Fixed amount per point.</summary>
    Fixed = 1,
    /// <summary>Percentage of each point's value.</summary>
    Percent = 2,
    /// <summary>Calculated standard error.</summary>
    StandardError = 4,
    /// <summary>Standard-deviation multiplier.</summary>
    StandardDeviation = -4155,
    /// <summary>Explicit worksheet ranges.</summary>
    Custom = -4114
}

/// <summary>Error-bar axis direction.</summary>
public enum ChartErrorBarDirection
{
    /// <summary>Value-axis error bars; horizontal on bar charts, vertical on column/line/scatter charts.</summary>
    Y = 1,
    /// <summary>X-value errors for XY scatter/bubble series.</summary>
    X = -4168
}

/// <summary>Error-bar signs to show.</summary>
public enum ChartErrorBarInclude
{
    /// <summary>Positive and negative bars.</summary>
    Both = 1,
    /// <summary>Positive bars only.</summary>
    Plus = 2,
    /// <summary>Negative bars only.</summary>
    Minus = 3
}

/// <summary>Native error-bar end caps.</summary>
public enum ChartErrorBarEndStyle
{
    /// <summary>Show end caps.</summary>
    Cap = 1,
    /// <summary>Hide end caps.</summary>
    NoCap = 2
}

/// <summary>Typed error-bar settings; nested JSON keys use camelCase.</summary>
public sealed class ChartErrorBarOptions
{
    /// <summary>False removes all error bars from the series.</summary>
    public bool Enabled { get; set; } = true;
    /// <summary>Native error-bar kind.</summary>
    public ChartErrorBarKind Kind { get; set; } = ChartErrorBarKind.Fixed;
    /// <summary>Native bar direction.</summary>
    public ChartErrorBarDirection Direction { get; set; } = ChartErrorBarDirection.Y;
    /// <summary>Which signs to show.</summary>
    public ChartErrorBarInclude Include { get; set; } = ChartErrorBarInclude.Both;
    /// <summary>Nonnegative fixed/percent amount or positive standard-deviation multiplier; omitted for StandardError/Custom.</summary>
    public double? Amount { get; set; }
    /// <summary>Custom ranges' worksheet; defaults to the chart worksheet.</summary>
    public string? SourceSheetName { get; set; }
    /// <summary>Custom positive values: one contiguous row/column with one nonnegative numeric value per point.</summary>
    public string? PlusRange { get; set; }
    /// <summary>Custom negative values, with the same geometry rules.</summary>
    public string? MinusRange { get; set; }
    /// <summary>Optional native cap style.</summary>
    public ChartErrorBarEndStyle? EndStyle { get; set; }
}

/// <summary>Native error-bar state, with explicit unavailable calculation getters.</summary>
public sealed class ChartErrorBarsResult : OperationResult
{
    /// <summary>Selected chart.</summary>
    public string ChartName { get; set; } = string.Empty;
    /// <summary>One-based selected series.</summary>
    public int SeriesIndex { get; set; }
    /// <summary>Native error-bar presence.</summary>
    public bool HasErrorBars { get; set; }
    /// <summary>False: Excel exposes no getters for calculation kind, direction, include, amount or custom range references.</summary>
    public bool SettingsReadable { get; set; }
    /// <summary>Explicit native read limitation, not cached request settings.</summary>
    public string ReadLimitations { get; set; } = "Excel COM does not expose error-bar kind, direction, include, amount or custom source getters.";
    /// <summary>Native end-cap setting when bars exist.</summary>
    public ChartErrorBarEndStyle? EndStyle { get; set; }
}

/// <summary>Selected-point material and supported marker changes.</summary>
public sealed class ChartPointOptions
{
    /// <summary>Solid fill as #RRGGBB.</summary>
    public string? FillColor { get; set; }
    /// <summary>Fill transparency, zero to one; not supported for line/scatter/radar marker points.</summary>
    public double? FillTransparency { get; set; }
    /// <summary>Outline as #RRGGBB.</summary>
    public string? LineColor { get; set; }
    /// <summary>Positive outline weight in points; not supported for line/scatter/radar marker points.</summary>
    public double? LineWeight { get; set; }
    /// <summary>Marker type for line/scatter/radar series; Picture requires an image and is unsupported here.</summary>
    public MarkerStyle? MarkerStyle { get; set; }
    /// <summary>Marker size, 2 to 72 points, for line/scatter/radar series.</summary>
    public int? MarkerSize { get; set; }
}

/// <summary>Native selected-point format.</summary>
public sealed class ChartPointFormatResult : OperationResult
{
    /// <summary>Selected chart.</summary>
    public string ChartName { get; set; } = string.Empty;
    /// <summary>One-based series index.</summary>
    public int SeriesIndex { get; set; }
    /// <summary>One-based point index.</summary>
    public int PointIndex { get; set; }
    /// <summary>Native fill color.</summary>
    public string? FillColor { get; set; }
    /// <summary>Whether Excel exposes an explicit RGB fill/marker background color.</summary>
    public bool FillColorAvailable { get; set; }
    /// <summary>Explicit automatic/mixed native color limitation.</summary>
    public string? FillColorReadError { get; set; }
    /// <summary>Native fill transparency when Excel returns a valid zero-to-one value.</summary>
    public double? FillTransparency { get; set; }
    /// <summary>Whether Excel exposes a valid point-transparency getter.</summary>
    public bool FillTransparencyAvailable { get; set; }
    /// <summary>Explicit native getter limitation when transparency cannot be inspected.</summary>
    public string? FillTransparencyReadError { get; set; }
    /// <summary>Native line color.</summary>
    public string? LineColor { get; set; }
    /// <summary>Whether Excel exposes an explicit RGB line/marker foreground color.</summary>
    public bool LineColorAvailable { get; set; }
    /// <summary>Explicit automatic/mixed native color limitation.</summary>
    public string? LineColorReadError { get; set; }
    /// <summary>Native line weight.</summary>
    public double? LineWeight { get; set; }
    /// <summary>Whether this point exposes a supported outline-weight getter.</summary>
    public bool LineWeightAvailable { get; set; }
    /// <summary>Whether this series has supported native marker getters.</summary>
    public bool MarkersSupported { get; set; }
    /// <summary>Native marker style when supported.</summary>
    public MarkerStyle? MarkerStyle { get; set; }
    /// <summary>Native marker size when supported.</summary>
    public int? MarkerSize { get; set; }
}

/// <summary>Native chart image filters.</summary>
public enum ChartImageFormat
{
    /// <summary>PNG image.</summary>
    Png,
    /// <summary>JPEG image.</summary>
    Jpeg,
    /// <summary>GIF image.</summary>
    Gif
}
