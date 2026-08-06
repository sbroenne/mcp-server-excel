namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal enum ExcelOperationState
{
    Queued = 0,
    Started = 1,
    Completed = 2,
    AbandonedBeforeStart = 3,
    OutcomeUnknown = 4
}

internal enum ExcelOperationInterruption
{
    NotStarted,
    StartedOutcomeUnknown,
    Completed
}

/// <summary>
/// Resolves the race between the caller's deadline and the single Excel STA worker.
/// A queued operation can be abandoned with certainty; once dispatch starts, a lost
/// caller must conservatively treat the mutation outcome as unknown.
/// </summary>
internal sealed class ExcelOperationLifecycle
{
    private int _state = (int)ExcelOperationState.Queued;

    public ExcelOperationState State => (ExcelOperationState)Volatile.Read(ref _state);

    public bool TryStart() =>
        Interlocked.CompareExchange(
            ref _state,
            (int)ExcelOperationState.Started,
            (int)ExcelOperationState.Queued) == (int)ExcelOperationState.Queued;

    public ExcelOperationInterruption Interrupt()
    {
        while (true)
        {
            var state = State;
            switch (state)
            {
                case ExcelOperationState.Queued:
                    if (Interlocked.CompareExchange(
                            ref _state,
                            (int)ExcelOperationState.AbandonedBeforeStart,
                            (int)ExcelOperationState.Queued) == (int)ExcelOperationState.Queued)
                    {
                        return ExcelOperationInterruption.NotStarted;
                    }
                    break;

                case ExcelOperationState.Started:
                    if (Interlocked.CompareExchange(
                            ref _state,
                            (int)ExcelOperationState.OutcomeUnknown,
                            (int)ExcelOperationState.Started) == (int)ExcelOperationState.Started)
                    {
                        return ExcelOperationInterruption.StartedOutcomeUnknown;
                    }
                    break;

                case ExcelOperationState.Completed:
                    return ExcelOperationInterruption.Completed;

                case ExcelOperationState.AbandonedBeforeStart:
                    return ExcelOperationInterruption.NotStarted;

                case ExcelOperationState.OutcomeUnknown:
                    return ExcelOperationInterruption.StartedOutcomeUnknown;

                default:
                    throw new InvalidOperationException($"Unknown Excel operation state: {state}.");
            }
        }
    }

    public void MarkCompleted()
    {
        while (true)
        {
            var state = State;
            switch (state)
            {
                case ExcelOperationState.Started:
                    if (Interlocked.CompareExchange(
                            ref _state,
                            (int)ExcelOperationState.Completed,
                            (int)ExcelOperationState.Started) == (int)ExcelOperationState.Started)
                    {
                        return;
                    }
                    break;

                case ExcelOperationState.Completed:
                case ExcelOperationState.OutcomeUnknown:
                    return;

                case ExcelOperationState.Queued:
                case ExcelOperationState.AbandonedBeforeStart:
                    throw new InvalidOperationException("An Excel operation cannot complete before it starts.");

                default:
                    throw new InvalidOperationException($"Unknown Excel operation state: {state}.");
            }
        }
    }
}
