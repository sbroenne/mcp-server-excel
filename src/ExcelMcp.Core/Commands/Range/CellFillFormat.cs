namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Native fill pattern and foreground/background colors.</summary>
/// <param name="Pattern">Native XlPattern code.</param>
/// <param name="Color">Fill color.</param>
/// <param name="PatternColor">Pattern color.</param>
/// <param name="Gradient">Complete gradient geometry and color stops, when present.</param>
public sealed record CellFillFormat(
    int? Pattern, CellColorFormat Color, CellColorFormat PatternColor,
    CellGradientFormat? Gradient);
