namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Native font properties; mixed rich-text fields are null.</summary>
/// <param name="Name">Font family.</param>
/// <param name="Size">Size in points.</param>
/// <param name="Bold">Bold state.</param>
/// <param name="Italic">Italic state.</param>
/// <param name="Underline">Native XlUnderlineStyle code.</param>
/// <param name="Strikethrough">Strikethrough state.</param>
/// <param name="Subscript">Subscript state.</param>
/// <param name="Superscript">Superscript state.</param>
/// <param name="ThemeFont">Native theme-font code.</param>
/// <param name="Color">Font color and theme information.</param>
public sealed record CellFontFormat(
    string? Name, double? Size, bool? Bold, bool? Italic, int? Underline,
    bool? Strikethrough, bool? Subscript, bool? Superscript, int? ThemeFont,
    CellColorFormat Color);
