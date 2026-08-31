namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Reports a rollback failure after the session manager attempted deterministic recovery.
/// </summary>
public sealed class WorkbookSavepointRollbackException : InvalidOperationException
{
    /// <summary>
    /// Initializes a rollback failure.
    /// </summary>
    public WorkbookSavepointRollbackException(
        string message,
        bool sessionRecovered,
        bool sessionClosed,
        string? recoveryFilePath,
        Exception rollbackException,
        Exception? recoveryException = null)
        : base(
            message,
            recoveryException == null
                ? rollbackException
                : new AggregateException(rollbackException, recoveryException))
    {
        SessionRecovered = sessionRecovered;
        SessionClosed = sessionClosed;
        RecoveryFilePath = recoveryFilePath;
    }

    /// <summary>Whether the pre-rollback workbook state was reopened under the same session ID.</summary>
    public bool SessionRecovered { get; }

    /// <summary>Whether the session had to be closed because recovery failed.</summary>
    public bool SessionClosed { get; }

    /// <summary>Caller-owned recovery file retained only when automatic recovery failed.</summary>
    public string? RecoveryFilePath { get; }
}
