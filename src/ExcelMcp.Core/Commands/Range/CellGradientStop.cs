namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>One native gradient color stop.</summary>
/// <param name="Position">Native relative stop position.</param>
/// <param name="Color">Resolved color, theme index, and tint.</param>
public sealed record CellGradientStop(double Position, CellColorFormat Color);
