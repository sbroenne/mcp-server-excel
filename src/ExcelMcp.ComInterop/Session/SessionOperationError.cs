namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Structured reason that a session cannot begin another operation.
/// </summary>
public enum SessionOperationError
{
    /// <summary>No error occurred.</summary>
    None,

    /// <summary>The caller omitted the session ID.</summary>
    MissingSessionId,

    /// <summary>The requested session does not exist.</summary>
    NotFound,

    /// <summary>The session is currently closing.</summary>
    Closing,

    /// <summary>The session is isolated after a failed close.</summary>
    Quarantined,

    /// <summary>A prior timeout or cancellation made the session unusable.</summary>
    TimedOutOrCancelled,

    /// <summary>The Excel process backing the session is no longer running.</summary>
    ExcelProcessDied
}
