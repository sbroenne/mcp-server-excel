namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Edge from which native directional filling proceeds.</summary>
public enum FillDirection
{
    /// <summary>Copy the top row down.</summary>
    Down,
    /// <summary>Copy the bottom row up.</summary>
    Up,
    /// <summary>Copy the rightmost column left.</summary>
    Left,
    /// <summary>Copy the leftmost column right.</summary>
    Right
}

/// <summary>Native Excel AutoFill behavior.</summary>
public enum AutoFillKind
{
    /// <summary>Excel infers the pattern.</summary>
    Default,
    /// <summary>Copy source cells.</summary>
    Copy,
    /// <summary>Extend the source series.</summary>
    Series,
    /// <summary>Fill formats only, leaving contents unchanged.</summary>
    Formats,
    /// <summary>Fill contents without formatting.</summary>
    WithoutFormatting,
    /// <summary>Extend dates by day.</summary>
    Days,
    /// <summary>Extend dates by weekday.</summary>
    Weekdays,
    /// <summary>Extend dates by month.</summary>
    Months,
    /// <summary>Extend dates by year.</summary>
    Years,
    /// <summary>Fit a linear trend.</summary>
    LinearTrend,
    /// <summary>Fit a growth trend.</summary>
    GrowthTrend,
    /// <summary>Use Excel Flash Fill where supported.</summary>
    FlashFill
}

/// <summary>Direction of each native series.</summary>
public enum SeriesOrientation
{
    /// <summary>Each row proceeds across columns.</summary>
    Rows,
    /// <summary>Each column proceeds down rows.</summary>
    Columns
}

/// <summary>Native DataSeries progression.</summary>
public enum SeriesKind
{
    /// <summary>Add the step.</summary>
    Linear,
    /// <summary>Multiply by the step.</summary>
    Growth,
    /// <summary>Advance dates by the selected date unit.</summary>
    Date,
    /// <summary>Let Excel extend source patterns.</summary>
    AutoFill
}

/// <summary>Date unit used by native date series.</summary>
public enum SeriesDateUnit
{
    /// <summary>Calendar days.</summary>
    Day,
    /// <summary>Weekdays.</summary>
    Weekday,
    /// <summary>Calendar months.</summary>
    Month,
    /// <summary>Calendar years.</summary>
    Year
}
