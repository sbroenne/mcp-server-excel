using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Native text parsing mode.</summary>
public enum TextParsingMode
{
    /// <summary>Split at enabled delimiters.</summary>
    Delimited,
    /// <summary>Split at explicit zero-based character positions.</summary>
    FixedWidth
}

/// <summary>Native quoted-field handling.</summary>
public enum TextFieldQualifier
{
    /// <summary>Double quotes surround fields.</summary>
    DoubleQuote,
    /// <summary>Single quotes surround fields.</summary>
    SingleQuote,
    /// <summary>No quoted-field handling.</summary>
    None
}

/// <summary>Excel's conversion for one parsed field.</summary>
public enum TextFieldType
{
    /// <summary>Excel infers numbers and dates.</summary>
    General = 1,
    /// <summary>Retain text, including leading zeroes.</summary>
    Text = 2,
    /// <summary>Month-day-year dates.</summary>
    Mdy = 3,
    /// <summary>Day-month-year dates.</summary>
    Dmy = 4,
    /// <summary>Year-month-day dates.</summary>
    Ymd = 5,
    /// <summary>Month-year-day dates.</summary>
    Myd = 6,
    /// <summary>Day-year-month dates.</summary>
    Dym = 7,
    /// <summary>Year-day-month dates.</summary>
    Ydm = 8,
    /// <summary>Do not output this field.</summary>
    Skip = 9,
    /// <summary>East Asian month-day-year dates.</summary>
    Emd = 10
}

/// <summary>Native FieldInfo entry; delimited indices are one-based, fixed-width positions zero-based.</summary>
[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class TextColumnField
{
    /// <summary>Delimited field index or fixed-width starting character position.</summary>
    public int Position { get; set; }
    /// <summary>Conversion for this field; General by default.</summary>
    public TextFieldType DataType { get; set; } = TextFieldType.General;
}

/// <summary>Native TextToColumns settings. Unknown settings are rejected.</summary>
[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class TextToColumnsOptions
{
    /// <summary>Delimited by default; FixedWidth requires explicit fields.</summary>
    public TextParsingMode Mode { get; set; }
    /// <summary>DoubleQuote by default.</summary>
    public TextFieldQualifier Qualifier { get; set; }
    /// <summary>Use tab delimiters.</summary>
    public bool Tab { get; set; }
    /// <summary>Use semicolon delimiters.</summary>
    public bool Semicolon { get; set; }
    /// <summary>Use comma delimiters.</summary>
    public bool Comma { get; set; }
    /// <summary>Use space delimiters.</summary>
    public bool Space { get; set; }
    /// <summary>Optional single-character delimiter.</summary>
    public string? OtherDelimiter { get; set; }
    /// <summary>Treat adjacent delimiters as one; false preserves empty fields.</summary>
    public bool ConsecutiveDelimiters { get; set; }
    /// <summary>Optional native conversion entries. FixedWidth requires an entry at zero.</summary>
    public List<TextColumnField>? Fields { get; set; }
    /// <summary>Optional decimal separator; omission uses Excel's current separator.</summary>
    public string? DecimalSeparator { get; set; }
    /// <summary>Optional thousands separator; omission uses Excel's current separator.</summary>
    public string? ThousandsSeparator { get; set; }
    /// <summary>Interpret a trailing minus sign as negative; true by default.</summary>
    public bool TrailingMinusNumbers { get; set; } = true;
}
