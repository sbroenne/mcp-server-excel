using System.Collections.Concurrent;
using System.IO.Pipes;
using StreamJsonRpc;

namespace Sbroenne.ExcelMcp.Service.Rpc;

internal sealed class DaemonHost : IDisposable
{
    internal static readonly TimeSpan InitialBackoff = TimeSpan.FromMilliseconds(100);
    internal static readonly TimeSpan MaxBackoff = TimeSpan.FromSeconds(5);

    private readonly ConcurrentDictionary<Task, byte> _activeConnectionTasks = new();
    private readonly CancellationTokenSource _shutdownCts = new();
    private readonly Func<ServiceRequest, Task<ServiceResponse>> _requestHandler;
    private readonly Func<int> _sessionCount;
    private readonly TimeProvider _timeProvider;
    private readonly Func<TimeSpan, CancellationToken, Task> _delay;
    private readonly Func<string, NamedPipeServerStream> _serverFactory;
    private readonly TaskCompletionSource _acceptLoopStopped =
        new(TaskCreationOptions.RunContinuationsAsynchronously);
    private DateTimeOffset _lastActivityTime;
    private bool _disposed;

    internal DaemonHost(
        Func<ServiceRequest, Task<ServiceResponse>> requestHandler,
        Func<int> sessionCount,
        TimeProvider? timeProvider = null,
        Func<TimeSpan, CancellationToken, Task>? delay = null,
        Func<string, NamedPipeServerStream>? serverFactory = null)
    {
        ArgumentNullException.ThrowIfNull(requestHandler);
        ArgumentNullException.ThrowIfNull(sessionCount);
        _requestHandler = requestHandler;
        _sessionCount = sessionCount;
        _timeProvider = timeProvider ?? TimeProvider.System;
        _delay = delay
            ?? ((duration, cancellationToken) =>
                Task.Delay(duration, _timeProvider, cancellationToken));
        _serverFactory = serverFactory ?? ServiceSecurity.CreateSecureServer;
        _lastActivityTime = _timeProvider.GetUtcNow();
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "The pipe accept-loop retries listener failures with bounded backoff so the daemon remains available; this boundary never converts a workbook operation failure into success.")]
    internal async Task RunAsync(string pipeName, TimeSpan? idleTimeout = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(pipeName);
        using var connectionLimit = new SemaphoreSlim(10, 10);

        if (idleTimeout.HasValue)
        {
            _ = Task.Run(
                () => MonitorIdleTimeoutAsync(idleTimeout.Value, _shutdownCts.Token),
                _shutdownCts.Token);
        }

        var currentBackoff = InitialBackoff;
        while (!_shutdownCts.IsCancellationRequested)
        {
            NamedPipeServerStream? server = null;
            try
            {
                server = _serverFactory(pipeName);
                await server.WaitForConnectionAsync(_shutdownCts.Token);
                currentBackoff = InitialBackoff;
                RecordActivity();

                var clientServer = server;
                server = null;
                var connectionTask = RunConnectionAsync(clientServer, connectionLimit);
                _activeConnectionTasks.TryAdd(connectionTask, 0);
                _ = connectionTask.ContinueWith(
                    completed => _activeConnectionTasks.TryRemove(completed, out _),
                    CancellationToken.None,
                    TaskContinuationOptions.ExecuteSynchronously,
                    TaskScheduler.Default);
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceWarning("Pipe accept failed: {0}", ex.Message);
                try
                {
                    await _delay(currentBackoff, _shutdownCts.Token);
                }
                catch (OperationCanceledException)
                {
                    break;
                }

                currentBackoff = TimeSpan.FromMilliseconds(
                    Math.Min(currentBackoff.TotalMilliseconds * 2, MaxBackoff.TotalMilliseconds));
            }
            finally
            {
                if (server != null)
                {
                    try
                    {
                        if (server.IsConnected)
                        {
                            server.Disconnect();
                        }
                    }
                    catch (IOException ex)
                    {
                        System.Diagnostics.Trace.TraceWarning("Pipe disconnect cleanup failed: {0}", ex.Message);
                    }

                    await server.DisposeAsync();
                }
            }
        }

        _acceptLoopStopped.TrySetResult();
        if (!_activeConnectionTasks.IsEmpty)
        {
            await Task.WhenAll(_activeConnectionTasks.Keys.Select(ObserveConnectionTaskAsync));
        }
    }

    internal void RequestShutdown() => _shutdownCts.Cancel();

    internal void RequestShutdownAfterResponse()
    {
        _ = ShutdownAfterResponseAsync();
    }

    internal void RecordActivity() => _lastActivityTime = _timeProvider.GetUtcNow();

    internal Task WaitForAcceptLoopStoppedAsync() => _acceptLoopStopped.Task;

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "A disconnected or faulted RPC connection is diagnosed and its pipe disposed without taking down unrelated clients. Request errors are transported by StreamJsonRpc before this connection boundary.")]
    private async Task RunConnectionAsync(
        NamedPipeServerStream clientServer,
        SemaphoreSlim connectionLimit)
    {
        await connectionLimit.WaitAsync();
        try
        {
            var rpcTarget = new DaemonRpcTarget(_requestHandler, RecordActivity);
            using var rpc = JsonRpc.Attach(clientServer, rpcTarget);
            await rpc.Completion;
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            System.Diagnostics.Debug.WriteLine($"RPC connection failed: {ex.Message}");
        }
        catch (OperationCanceledException ex)
        {
            System.Diagnostics.Debug.WriteLine($"RPC connection cancelled: {ex.Message}");
        }
        finally
        {
            connectionLimit.Release();
            try
            {
                if (clientServer.IsConnected)
                {
                    clientServer.Disconnect();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Pipe disconnect cleanup failed: {ex.Message}");
            }

            try
            {
                await clientServer.DisposeAsync();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Pipe disposal cleanup failed: {ex.Message}");
            }
        }
    }

    private async Task MonitorIdleTimeoutAsync(
        TimeSpan idleTimeout,
        CancellationToken cancellationToken)
    {
        try
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                await _delay(TimeSpan.FromSeconds(30), cancellationToken);
                if (_sessionCount() > 0)
                {
                    RecordActivity();
                    continue;
                }

                if (_timeProvider.GetUtcNow() - _lastActivityTime >= idleTimeout)
                {
                    RequestShutdown();
                    break;
                }
            }
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            return;
        }
    }

    private async Task ShutdownAfterResponseAsync()
    {
        try
        {
            await _delay(TimeSpan.FromMilliseconds(100), _shutdownCts.Token);
            RequestShutdown();
        }
        catch (OperationCanceledException) when (_shutdownCts.IsCancellationRequested)
        {
            return;
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Shutdown observes and diagnoses already-faulted connection tasks while allowing the remaining connections to drain; no operation result is changed.")]
    private static async Task ObserveConnectionTaskAsync(Task connectionTask)
    {
        try
        {
            await connectionTask;
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            System.Diagnostics.Debug.WriteLine($"RPC connection drain failed: {ex.Message}");
        }
        catch (OperationCanceledException ex)
        {
            System.Diagnostics.Debug.WriteLine($"RPC connection drain cancelled: {ex.Message}");
        }
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        _shutdownCts.Cancel();
        _shutdownCts.Dispose();
    }
}
