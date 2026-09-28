namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

internal readonly struct OperationDeadline
{
    private readonly long _startedAt;
    private readonly TimeSpan _timeout;
    private readonly TimeProvider _timeProvider;

    private OperationDeadline(TimeSpan timeout, TimeProvider timeProvider)
    {
        ArgumentNullException.ThrowIfNull(timeProvider);
        _timeProvider = timeProvider;
        _startedAt = timeProvider.GetTimestamp();
        _timeout = timeout;
    }

    internal static OperationDeadline Start(TimeSpan timeout) => new(timeout, TimeProvider.System);

    internal static OperationDeadline Start(TimeSpan timeout, TimeProvider timeProvider) => new(timeout, timeProvider);

    internal TimeSpan Remaining
    {
        get
        {
            if (_timeout <= TimeSpan.Zero)
                return TimeSpan.Zero;

            var remaining = _timeout - _timeProvider.GetElapsedTime(_startedAt);
            return remaining > TimeSpan.Zero ? remaining : TimeSpan.Zero;
        }
    }

    internal bool IsExpired => Remaining <= TimeSpan.Zero;

    internal TimeSpan Cap(TimeSpan maximum)
    {
        var remaining = Remaining;
        return remaining <= maximum ? remaining : maximum;
    }
}
