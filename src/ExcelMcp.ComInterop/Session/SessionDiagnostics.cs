using System.Diagnostics;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Diagnostic tracing for Excel COM session operations.
/// </summary>
public static class SessionDiagnostics
{
    private const string DiagnosticsEnvVar = "EXCELMCP_DIAGNOSTICS";
    private static readonly AsyncLocal<CloseTimingCollection?> CloseTimings = new();

    internal static CloseTimingCollection? CurrentCloseTimings => CloseTimings.Value;

    internal static CloseTimingCollection BeginCloseTimings() => new();

    internal static IDisposable? UseCloseTimings(CloseTimingCollection? timings) =>
        timings is null ? null : new CloseTimingContext(timings);

    /// <summary>
    /// Gets whether diagnostic tracing is enabled via EXCELMCP_DIAGNOSTICS environment variable.
    /// </summary>
    public static bool IsEnabled =>
        string.Equals(Environment.GetEnvironmentVariable(DiagnosticsEnvVar), "1", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(Environment.GetEnvironmentVariable(DiagnosticsEnvVar), "true", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Writes a diagnostic message to stderr if diagnostics are enabled.
    /// </summary>
    /// <param name="message">Message to write</param>
    public static void WriteStdErr(string message)
    {
        if (IsEnabled)
        {
            Console.Error.WriteLine(message);
        }
    }

    internal static T MeasureCloseStage<T>(string stage, Func<T> operation)
    {
        using var timing = BeginCloseStage(stage);
        return operation();
    }

    internal static void MeasureCloseStage(string stage, Action operation)
    {
        using var timing = BeginCloseStage(stage);
        operation();
    }

    internal static CloseStageTiming? BeginCloseStage(string stage) =>
        CloseTimings.Value is { } timings ? new CloseStageTiming(stage, timings) : null;

    internal sealed class CloseTimingCollection : IDisposable
    {
        private readonly CloseTimingCollection? _previous = CloseTimings.Value;
        private readonly Lock _lock = new();
        private readonly Dictionary<string, double> _elapsed = new(StringComparer.Ordinal);

        internal CloseTimingCollection() => CloseTimings.Value = this;

        internal void Record(string stage, double elapsedMs)
        {
            lock (_lock)
                _elapsed[stage] = _elapsed.GetValueOrDefault(stage) + elapsedMs;
        }

        internal Dictionary<string, double> Snapshot()
        {
            lock (_lock)
                return _elapsed.ToDictionary(entry => entry.Key, entry => Math.Round(entry.Value, 2));
        }

        public void Dispose() => CloseTimings.Value = _previous;
    }

    private sealed class CloseTimingContext(CloseTimingCollection timings) : IDisposable
    {
        private readonly CloseTimingCollection? _previous = SetCurrent(timings);

        private static CloseTimingCollection? SetCurrent(CloseTimingCollection timings)
        {
            var previous = CloseTimings.Value;
            CloseTimings.Value = timings;
            return previous;
        }

        public void Dispose() => CloseTimings.Value = _previous;
    }

    internal sealed class CloseStageTiming(string stage, CloseTimingCollection timings) : IDisposable
    {
        private readonly long _started = Stopwatch.GetTimestamp();

        public void Dispose() => timings.Record(stage, Stopwatch.GetElapsedTime(_started).TotalMilliseconds);
    }
}
