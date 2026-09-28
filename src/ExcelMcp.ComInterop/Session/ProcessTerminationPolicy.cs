namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal static class ProcessTerminationPolicy
{
    internal static readonly TimeSpan ProcessExitTimeout = TimeSpan.FromSeconds(10);
    internal static readonly TimeSpan NormalGraceTimeout = TimeSpan.FromSeconds(10);
    internal static readonly TimeSpan NormalForcedExitTimeout = TimeSpan.FromSeconds(5);
    internal static readonly TimeSpan NormalShutdownBudget = TimeSpan.FromSeconds(15);

    internal static async Task<bool> TryCompleteAsync(
        TimeSpan waitBeforeTermination,
        TimeSpan waitAfterTermination,
        Func<TimeSpan, CancellationToken, Task<ProcessWaitOutcome>> waitForExitAsync,
        Func<ProcessTerminationOutcome> requestTermination,
        CancellationToken cancellationToken,
        Action<bool> setTerminated,
        TimeSpan? overallTimeout = null,
        TimeProvider? timeProvider = null)
    {
        ArgumentNullException.ThrowIfNull(waitForExitAsync);
        ArgumentNullException.ThrowIfNull(requestTermination);
        ArgumentNullException.ThrowIfNull(setTerminated);
        var clock = timeProvider ?? TimeProvider.System;
        var started = clock.GetTimestamp();

        var initialWait = await waitForExitAsync(
            LimitWait(waitBeforeTermination),
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
        return await waitForExitAsync(LimitWait(waitAfterTermination), cancellationToken)
            == ProcessWaitOutcome.Exited;

        TimeSpan LimitWait(TimeSpan requested)
        {
            if (overallTimeout is not { } budget) return requested;
            var remaining = budget - clock.GetElapsedTime(started);
            return remaining <= TimeSpan.Zero
                ? TimeSpan.Zero
                : remaining < requested ? remaining : requested;
        }
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
