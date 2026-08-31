namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal sealed class WorkbookRestoreException : InvalidOperationException
{
    internal WorkbookRestoreException(
        bool sessionRecovered,
        Exception rollbackException,
        Exception? recoveryException = null)
        : base(
            sessionRecovered
                ? "Workbook rollback failed, but the pre-rollback state was recovered."
                : "Workbook rollback and emergency recovery both failed.",
            recoveryException == null
                ? rollbackException
                : new AggregateException(rollbackException, recoveryException))
    {
        SessionRecovered = sessionRecovered;
        RollbackException = rollbackException;
        RecoveryException = recoveryException;
    }

    internal bool SessionRecovered { get; }

    internal Exception RollbackException { get; }

    internal Exception? RecoveryException { get; }
}
