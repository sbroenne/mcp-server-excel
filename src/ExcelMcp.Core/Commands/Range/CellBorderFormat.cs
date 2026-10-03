namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>One native edge, diagonal, or inside border.</summary>
/// <param name="Edge">Native XlBordersIndex name.</param>
/// <param name="LineStyle">Native XlLineStyle code.</param>
/// <param name="Weight">Native XlBorderWeight code.</param>
/// <param name="Color">Border color and theme information.</param>
public sealed record CellBorderFormat(
    string Edge, int? LineStyle, int? Weight, CellColorFormat Color);
