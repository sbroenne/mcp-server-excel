using System.Diagnostics;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Diagnostic tracing for Excel COM session operations.
/// </summary>
public static class SessionDiagnostics
{
    private const string DiagnosticsEnvVar = "EXCELMCP_DIAGNOSTICS";
    private const long MaximumCloseLogBytes = 1024 * 1024;
    private static readonly Lock CloseLogLock = new();

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
        var result = operation();
        timing?.Complete();
        return result;
    }

    internal static void MeasureCloseStage(string stage, Action operation)
    {
        using var timing = BeginCloseStage(stage);
        operation();
        timing?.Complete();
    }

    internal static async Task MeasureCloseStageAsync(string stage, Func<Task> operation)
    {
        using var timing = BeginCloseStage(stage);
        await operation();
        timing?.Complete();
    }

    internal static CloseStageTiming? BeginCloseStage(string stage) =>
        IsEnabled ? new CloseStageTiming(stage) : null;

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Opt-in local timing is a diagnostic side channel. File failures are reported through Trace and must not replace a workbook result or its original exception.")]
    private static void WriteCloseTiming(
        string stage, string stageId, string eventName, double? elapsedMs, bool? completed)
    {
        var record = JsonSerializer.Serialize(new
        {
            timestampUtc = DateTime.UtcNow,
            processId = Environment.ProcessId,
            stageId,
            stage,
            @event = eventName,
            elapsedMs,
            completed
        });
        try
        {
            var localDirectory = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            if (string.IsNullOrEmpty(localDirectory))
                throw new IOException("The current user's local application-data directory is unavailable.");
            var directory = Path.Combine(localDirectory, "ExcelMcp", "diagnostics");
            var path = Path.Combine(directory, $"close-{Environment.ProcessId}.jsonl");
            lock (CloseLogLock)
            {
                Directory.CreateDirectory(directory);
                if (File.Exists(path) && new FileInfo(path).Length >= MaximumCloseLogBytes)
                    File.Move(path, path + ".previous", overwrite: true);
                File.AppendAllText(path, record + Environment.NewLine);
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            Trace.TraceWarning("Could not write the local Excel close timing log: {0}", ex);
        }
    }

    internal sealed class CloseStageTiming : IDisposable
    {
        private readonly string _stage;
        private readonly string _stageId = Guid.NewGuid().ToString("N");
        private readonly long _started;
        private bool _completed;

        internal CloseStageTiming(string stage)
        {
            _stage = stage;
            WriteCloseTiming(_stage, _stageId, "started", null, null);
            _started = Stopwatch.GetTimestamp();
        }

        internal void Complete() => _completed = true;

        public void Dispose() => WriteCloseTiming(
            _stage, _stageId, "finished", Stopwatch.GetElapsedTime(_started).TotalMilliseconds, _completed);
    }
}
