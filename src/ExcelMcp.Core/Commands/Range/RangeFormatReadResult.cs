using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Complete formatting inspection for an explicitly requested scope.</summary>
public sealed class RangeFormatReadResult : ResultBase
{
    /// <summary>The resolved worksheet.</summary>
    public string SheetName { get; set; } = string.Empty;
    /// <summary>The resolved absolute requested scope.</summary>
    public string RangeAddress { get; set; } = string.Empty;
    /// <summary>The requested formatting view.</summary>
    public FormatView View { get; set; }
    /// <summary>The exact number of inspected cells.</summary>
    public long CellCount { get; set; }
    /// <summary>Every requested cell, ordered by row and column.</summary>
    public List<CellFormatRead> Cells { get; set; } = [];
}
