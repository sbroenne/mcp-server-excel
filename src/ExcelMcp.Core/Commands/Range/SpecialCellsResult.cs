using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Complete native cell discovery for an explicitly requested scope.</summary>
public sealed class SpecialCellsResult : ResultBase
{
    /// <summary>The resolved worksheet name, including for named-range requests.</summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>The resolved requested scope, as an absolute A1 address.</summary>
    public string RangeAddress { get; set; } = string.Empty;

    /// <summary>The requested cell category.</summary>
    public SpecialCellKind CellKind { get; set; }

    /// <summary>All matching rectangular areas as absolute A1 addresses, ordered by position.</summary>
    public List<string> Areas { get; set; } = [];

    /// <summary>The exact total number of matching cells, not the number of areas.</summary>
    public long CellCount { get; set; }
}
