using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>Observed native cache state after an explicit cache-sharing request.</summary>
public sealed class PivotCacheSharingResult : OperationResult
{
    /// <summary>Worksheet containing the moved report.</summary>
    public string SheetName { get; set; } = string.Empty;
    /// <summary>Name of the moved report.</summary>
    public string PivotTableName { get; set; } = string.Empty;
    /// <summary>Worksheet containing the chosen cache owner.</summary>
    public string ShareCacheWithSheet { get; set; } = string.Empty;
    /// <summary>Name of the chosen cache owner.</summary>
    public string ShareCacheWith { get; set; } = string.Empty;
    /// <summary>Selected report's native cache index before the request.</summary>
    public int PreviousCacheIndex { get; set; }
    /// <summary>Selected report's actual native cache index after the request.</summary>
    public int CacheIndex { get; set; }
    /// <summary>Chosen report's actual native cache index.</summary>
    public int TargetCacheIndex { get; set; }
    /// <summary>False when the reports already shared their cache.</summary>
    public bool Changed { get; set; }
    /// <summary>Actual users of the selected report's current cache, including itself.</summary>
    public List<PivotTableIdentity> SharedPivotTables { get; set; } = [];
    /// <summary>Reports still using the original cache, empty when none remain.</summary>
    public List<PivotTableIdentity> PreviousCacheUsers { get; set; } = [];
}
