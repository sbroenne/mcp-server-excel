namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Native color and theme information; null values are unset or mixed.</summary>
/// <param name="Rgb">Resolved #RRGGBB color, when present and uniform.</param>
/// <param name="ColorIndex">Native Excel palette/automatic/none index.</param>
/// <param name="ThemeColor">Native theme index, or null for a non-theme color.</param>
/// <param name="TintAndShade">Native brightness adjustment.</param>
public sealed record CellColorFormat(
    string? Rgb, int? ColorIndex, int? ThemeColor, double? TintAndShade);
