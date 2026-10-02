namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Selected native formatting snapshots for one cell.</summary>
/// <param name="Address">Absolute A1 cell address.</param>
/// <param name="Row">One-based worksheet row.</param>
/// <param name="Column">One-based worksheet column.</param>
/// <param name="Stored">Stored snapshot when requested.</param>
/// <param name="Displayed">Displayed snapshot when requested.</param>
public sealed record CellFormatRead(
    string Address, int Row, int Column,
    CellFormatSnapshot? Stored, CellFormatSnapshot? Displayed);
