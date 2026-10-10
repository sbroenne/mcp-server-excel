using Microsoft.ApplicationInsights.DataContracts;
using Sbroenne.ExcelMcp.CLI.Infrastructure;

namespace Sbroenne.ExcelMcp.CLI.Telemetry;

/// <summary>
/// Creates the real sink on a background task so building the telemetry SDK
/// overlaps the command's own work. A command that never tracks anything
/// (help, version, the background service) never waits for it. The initialization
/// timeout is one budget counted from <see cref="Start"/>, shared by every call.
/// </summary>
internal sealed class DeferredTelemetrySink(
    Func<ICliTelemetrySink?> factory,
    TimeSpan initializationTimeout,
    TimeSpan shutdownTimeout)
{
    private readonly object _stateLock = new();
    private Task<ICliTelemetrySink?>? _initialization;
    private OperationDeadline _initializationDeadline;
    private bool _shutdown;
    private int _tracked;

    /// <summary>Begins creating the sink. Safe to call repeatedly and from any thread.</summary>
    internal void Start() => _ = EnsureStarted();

    /// <summary>
    /// Forwards the items once the sink exists. Waits only for what is left of the
    /// initialization timeout, and drops the items when creation is still running after it.
    /// </summary>
    internal void Track(EventTelemetry eventTelemetry, RequestTelemetry requestTelemetry)
    {
        try
        {
            var (initialization, deadline) = EnsureStarted();
            if (!initialization.Wait(deadline.Remaining)
                || initialization.Result is not { } sink)
            {
                return;
            }

            // Recorded before forwarding: a sink that fails part-way may still hold data to flush.
            Volatile.Write(ref _tracked, 1);
            sink.Track(eventTelemetry, requestTelemetry);
        }
        catch (Exception)
        {
        }
    }

    /// <summary>
    /// Flushes and disposes the sink, waiting at most the shutdown timeout. Returns at
    /// once when nothing was tracked, even if creation is still running.
    /// </summary>
    internal void Shutdown()
    {
        Task<ICliTelemetrySink?>? initialization;
        lock (_stateLock)
        {
            if (_shutdown)
            {
                return;
            }

            _shutdown = true;
            initialization = _initialization;
        }

        if (initialization == null || Volatile.Read(ref _tracked) == 0)
        {
            return;
        }

        try
        {
            // Tracking implies the sink was created, so this task is already complete.
            if (initialization.Result is { } sink)
            {
                Task.Run(() => FlushAndDisposeAsync(sink)).Wait(shutdownTimeout);
            }
        }
        catch (Exception)
        {
        }
    }

    private (Task<ICliTelemetrySink?> Initialization, OperationDeadline Deadline) EnsureStarted()
    {
        lock (_stateLock)
        {
            if (_initialization == null)
            {
                _initializationDeadline = OperationDeadline.Start(initializationTimeout);
                _initialization = Task.Run(CreateSink);
            }

            return (_initialization, _initializationDeadline);
        }
    }

    private ICliTelemetrySink? CreateSink()
    {
        try
        {
            return factory();
        }
        catch (Exception)
        {
            return null;
        }
    }

    private static async Task FlushAndDisposeAsync(ICliTelemetrySink sink)
    {
        try
        {
            await sink.FlushAsync();
        }
        catch (Exception)
        {
        }
        finally
        {
            try
            {
                sink.Dispose();
            }
            catch (Exception)
            {
            }
        }
    }
}
