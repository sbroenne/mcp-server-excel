namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// The operation deadline expired while the request was still waiting for the Excel STA worker.
/// The delegate was never dispatched, so retrying is safe and the session remains usable.
/// </summary>
public sealed class ExcelOperationNotStartedTimeoutException : TimeoutException
{
    /// <summary>Creates a known-not-started timeout failure.</summary>
    /// <param name="message">Human-readable timeout context.</param>
    public ExcelOperationNotStartedTimeoutException(string message)
        : base(message)
    {
    }
}

/// <summary>
/// The caller cancelled while the request was still waiting for the Excel STA worker.
/// The delegate was never dispatched, so retrying is safe and the session remains usable.
/// </summary>
public sealed class ExcelOperationNotStartedCanceledException : OperationCanceledException
{
    /// <summary>Creates a known-not-started cancellation failure.</summary>
    /// <param name="message">Human-readable cancellation context.</param>
    /// <param name="cancellationToken">The caller token that cancelled admission or dispatch.</param>
    public ExcelOperationNotStartedCanceledException(string message, CancellationToken cancellationToken)
        : base(message, cancellationToken)
    {
    }
}
