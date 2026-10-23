namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>Bounded execution metadata; never contains a workbook result.</summary>
public sealed record ExcelOperationStatus(
    string State,
    double ElapsedSeconds,
    bool WaitExpired,
    bool ExecutionPending,
    string? ErrorMessage,
    string? WindowReveal)
{
    /// <summary>The Service command whose Excel callback owns this status.</summary>
    public string? Command { get; init; }
    /// <summary>When the currently reported operation started, in UTC.</summary>
    public DateTime StartedAtUtc { get; init; }
}

/// <summary>A startup wait expired, but its session still owns the Excel instance.</summary>
public sealed class SessionStartupPendingException(string sessionId, string message, Exception cause)
    : TimeoutException(message, cause)
{
    /// <summary>The retained session, including an unfinished startup.</summary>
    public string SessionId { get; } = sessionId;
}

internal sealed class ExcelStartupPendingException(string message) : TimeoutException(message);

/// <summary>A failed lifecycle operation retains its session and exact failure cause.</summary>
public sealed class SessionRetainedException(string sessionId, string message, Exception cause)
    : InvalidOperationException(message, cause)
{
    /// <summary>The session whose workbook or teardown is still owned.</summary>
    public string SessionId { get; } = sessionId;
}
