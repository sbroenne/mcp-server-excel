namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Refresh status for a connection or query table.
/// </summary>
public sealed class RefreshStatusResult : ResultBase
{
    /// <summary>Whether the COM object exposes refresh status.</summary>
    public bool SupportsRefreshStatus { get; init; }

    /// <summary>Whether a refresh is currently active.</summary>
    public bool IsRefreshing { get; init; }
}
