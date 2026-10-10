using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>A worksheet-qualified PivotTable identity.</summary>
public sealed class PivotTableIdentity
{
    /// <summary>Worksheet containing the PivotTable.</summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>PivotTable name on that worksheet.</summary>
    public string PivotTableName { get; set; } = string.Empty;
}

/// <summary>Native connection state without connection strings or account information.</summary>
public sealed class PivotConnectionResult : OperationResult
{
    /// <summary>Worksheet containing the selected PivotTable.</summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>Selected PivotTable name.</summary>
    public string PivotTableName { get; set; } = string.Empty;

    /// <summary>Workbook connection name; absent for worksheet-backed sources.</summary>
    public string? ConnectionName { get; set; }

    /// <summary>Current native cache index.</summary>
    public int CacheIndex { get; set; }

    /// <summary>Whether this is an external source rather than worksheet data or the workbook Data Model.</summary>
    public bool IsExternal { get; set; }

    /// <summary>Whether the cache uses OLAP fields and measures.</summary>
    public bool IsOlap { get; set; }

    /// <summary>Whether the source is the workbook's internal Data Model.</summary>
    public bool IsDataModel { get; set; }

    /// <summary>All PivotTables sharing the selected cache, including the selected table.</summary>
    public List<PivotTableIdentity> SharedPivotTables { get; set; } = [];

    /// <summary>Slicer/timeline caches connected to this exact worksheet-qualified PivotTable.</summary>
    public List<string> ConnectedSlicerCaches { get; set; } = [];
}
