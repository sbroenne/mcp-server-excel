namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal static class ProcessTerminationPolicy
{
    internal static readonly TimeSpan ProcessExitTimeout = TimeSpan.FromSeconds(10);

    internal static async Task<bool> TryCompleteAsync(
        TimeSpan waitBeforeTermination,
        TimeSpan waitAfterTermination,
        Func<TimeSpan, CancellationToken, Task<ProcessWaitOutcome>> waitForExitAsync,
        Func<ProcessTerminationOutcome> requestTermination,
        CancellationToken cancellationToken,
        Action<bool> setTerminated)
    {
        ArgumentNullException.ThrowIfNull(waitForExitAsync);
        ArgumentNullException.ThrowIfNull(requestTermination);
        ArgumentNullException.ThrowIfNull(setTerminated);

        var initialWait = await waitForExitAsync(
            waitBeforeTermination,
            cancellationToken);
        if (initialWait == ProcessWaitOutcome.Exited)
        {
            return true;
        }

        if (initialWait == ProcessWaitOutcome.Failed)
        {
            return false;
        }

        var termination = requestTermination();
        if (termination == ProcessTerminationOutcome.ConfirmedExited)
        {
            return true;
        }

        setTerminated(termination == ProcessTerminationOutcome.Requested);
        return await waitForExitAsync(waitAfterTermination, cancellationToken)
            == ProcessWaitOutcome.Exited;
    }

    internal enum ProcessWaitOutcome
    {
        Exited,
        TimedOut,
        Failed
    }

    internal enum ProcessTerminationOutcome
    {
        Requested,
        ConfirmedExited,
        Unavailable
    }
}
