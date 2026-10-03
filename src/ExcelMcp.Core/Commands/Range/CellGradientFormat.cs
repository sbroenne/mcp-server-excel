namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Native linear or rectangular gradient geometry.</summary>
/// <param name="Kind">linear or rectangular.</param>
/// <param name="Degree">Linear gradient angle.</param>
/// <param name="Top">Rectangular gradient top position.</param>
/// <param name="Bottom">Rectangular gradient bottom position.</param>
/// <param name="Left">Rectangular gradient left position.</param>
/// <param name="Right">Rectangular gradient right position.</param>
/// <param name="Stops">All native color stops in position order.</param>
public sealed record CellGradientFormat(
    string Kind, double? Degree, double? Top, double? Bottom, double? Left,
    double? Right, List<CellGradientStop> Stops);
