namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Native spill relationships for one requested cell.</summary>
/// <param name="Address">Absolute inspected address.</param>
/// <param name="Row">Worksheet row.</param>
/// <param name="Column">Worksheet column.</param>
/// <param name="State">Ordinary, source, result, or blocked.</param>
/// <param name="SourceAddress">Native source address, including a blocked formula itself.</param>
/// <param name="SourceFormula">The source formula in dynamic-array A1 notation.</param>
/// <param name="SpillAddress">Established native result extent, absent for blocked formulas.</param>
/// <param name="SpillRows">Established result row count.</param>
/// <param name="SpillColumns">Established result column count.</param>
public sealed record SpillCellInfo(
    string Address, int Row, int Column, SpillCellState State,
    string? SourceAddress = null, string? SourceFormula = null,
    string? SpillAddress = null, int? SpillRows = null, int? SpillColumns = null);
