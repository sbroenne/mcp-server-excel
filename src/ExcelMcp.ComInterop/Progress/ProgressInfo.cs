namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>
/// Domain-agnostic progress information for long-running operations.
/// Used with <see cref="IProgress{T}"/> to report progress from Core commands.
/// </summary>
public sealed class ProgressInfo
{
    /// <summary>Current step number (e.g., 3 of 10).</summary>
    public required float Current { get; init; }

    /// <summary>Total number of steps, if known.</summary>
    public float? Total { get; init; }

    /// <summary>Human-readable status message (e.g., query name being refreshed).</summary>
    public string? Message { get; init; }
}
