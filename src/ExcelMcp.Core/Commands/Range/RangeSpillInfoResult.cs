using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Complete native spill inspection; unavailable capability fails explicitly.</summary>
public sealed class RangeSpillInfoResult : ResultBase
{
    /// <summary>The resolved worksheet.</summary>
    public string SheetName { get; set; } = string.Empty;
    /// <summary>The resolved absolute requested scope.</summary>
    public string RangeAddress { get; set; } = string.Empty;
    /// <summary>Supported native inspection; unsupported sessions return an error instead.</summary>
    public string Capability { get; } = "supported";
    /// <summary>The exact number of inspected cells.</summary>
    public long CellCount { get; set; }
    /// <summary>Every requested cell, ordered by row and column.</summary>
    public List<SpillCellInfo> Cells { get; set; } = [];
}
