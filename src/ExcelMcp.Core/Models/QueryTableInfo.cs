namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Summary information for a worksheet QueryTable.
/// </summary>
public sealed class QueryTableInfo
{
    /// <summary>QueryTable name.</summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>Worksheet containing the QueryTable.</summary>
    public string SheetName { get; init; } = string.Empty;

    /// <summary>Top-left destination address.</summary>
    public string Destination { get; init; } = string.Empty;

    /// <summary>Source type: text, web, database, or other.</summary>
    public string SourceType { get; init; } = string.Empty;

    /// <summary>Whether the QueryTable is currently refreshing.</summary>
    public bool IsRefreshing { get; init; }
}
