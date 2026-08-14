namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result containing worksheet QueryTables.
/// </summary>
public sealed class QueryTableListResult : ResultBase
{
    /// <summary>QueryTables in workbook worksheet order.</summary>
    public List<QueryTableInfo> QueryTables { get; init; } = [];
}
