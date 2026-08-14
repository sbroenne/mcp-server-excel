namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result of requesting cancellation for a connection or query table refresh.
/// </summary>
public sealed class RefreshCancellationResult : ResultBase
{
    /// <summary>Whether the COM object exposes refresh cancellation.</summary>
    public bool SupportsCancellation { get; init; }

    /// <summary>Whether a refresh was active when cancellation was requested.</summary>
    public bool WasRefreshing { get; init; }

    /// <summary>
    /// Whether a cancellation request was issued while refresh was active.
    /// Excel COM does not synchronously confirm that provider cancellation completed.
    /// </summary>
    public bool Cancelled { get; init; }
}
