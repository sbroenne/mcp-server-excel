using System.Diagnostics;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

internal readonly struct OperationDeadline
{
    private readonly long _startedAt;
    private readonly TimeSpan _timeout;

    private OperationDeadline(TimeSpan timeout)
    {
        _startedAt = Stopwatch.GetTimestamp();
        _timeout = timeout;
    }

    internal static OperationDeadline Start(TimeSpan timeout) => new(timeout);

    internal TimeSpan Remaining
    {
        get
        {
            var remaining = _timeout - Stopwatch.GetElapsedTime(_startedAt);
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
