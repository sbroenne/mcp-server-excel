using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Complete per-cell native protection flags for the exact scope.</summary>
public sealed class RangeCellProtectionResult : OperationResult
{
    /// <summary>Resolved worksheet name.</summary>
    public string SheetName { get; set; } = string.Empty;
    /// <summary>Resolved exact scope.</summary>
    public string RangeAddress { get; set; } = string.Empty;
    /// <summary>Every requested cell, with no preview limit.</summary>
    public List<CellProtectionRead> Cells { get; set; } = [];
}

/// <summary>Native protection flags of one cell.</summary>
public sealed record CellProtectionRead(string Address, bool Locked, bool FormulaHidden);
