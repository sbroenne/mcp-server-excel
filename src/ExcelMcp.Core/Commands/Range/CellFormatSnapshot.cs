namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Native formatting for one cell, with explicit mixed-property paths.</summary>
public sealed class CellFormatSnapshot
{
    /// <summary>Font properties.</summary>
    public required CellFontFormat Font { get; init; }
    /// <summary>Fill properties.</summary>
    public required CellFillFormat Fill { get; init; }
    /// <summary>All eight native border positions.</summary>
    public required List<CellBorderFormat> Borders { get; init; }
    /// <summary>Invariant Excel number format.</summary>
    public string? NumberFormat { get; init; }
    /// <summary>Native horizontal alignment code.</summary>
    public int? HorizontalAlignment { get; init; }
    /// <summary>Native vertical alignment code.</summary>
    public int? VerticalAlignment { get; init; }
    /// <summary>Text wrapping.</summary>
    public bool? WrapText { get; init; }
    /// <summary>Text shrinking to fit.</summary>
    public bool? ShrinkToFit { get; init; }
    /// <summary>Automatic indentation.</summary>
    public bool? AddIndent { get; init; }
    /// <summary>Indentation level.</summary>
    public int? IndentLevel { get; init; }
    /// <summary>Native text rotation/vertical orientation.</summary>
    public int? Orientation { get; init; }
    /// <summary>Native reading-order code.</summary>
    public int? ReadingOrder { get; init; }
    /// <summary>Cell locking state; effective only with sheet protection.</summary>
    public bool? Locked { get; init; }
    /// <summary>Formula-hiding state; effective only with sheet protection.</summary>
    public bool? FormulaHidden { get; init; }
    /// <summary>Merge state.</summary>
    public bool? MergeCells { get; init; }
    /// <summary>Applied cell style.</summary>
    public string? StyleName { get; init; }
    /// <summary>Column width in Excel character units.</summary>
    public double? ColumnWidth { get; init; }
    /// <summary>Row height in points.</summary>
    public double? RowHeight { get; init; }
    /// <summary>Property paths for which Excel reports mixed values.</summary>
    public required List<string> MixedFields { get; init; }
}
