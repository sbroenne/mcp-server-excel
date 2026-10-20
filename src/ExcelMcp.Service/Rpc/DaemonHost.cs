using System.Collections.Concurrent;
using System.IO.Pipes;
using StreamJsonRpc;

namespace Sbroenne.ExcelMcp.Service.Rpc;

internal sealed class DaemonHost : IDisposable
{
    internal static readonly TimeSpan InitialBackoff = TimeSpan.FromMilliseconds(100);
    internal static readonly TimeSpan MaxBackoff = TimeSpan.FromSeconds(5);
    internal const int MaximumConnections = 64;

    private readonly ConcurrentDictionary<Task, byte> _activeConnectionTasks = new();
    private readonly ConcurrentDictionary<string, byte> _busySessions = new(StringComparer.Ordinal);
    private readonly CancellationTokenSource _shutdownCts = new();
    private readonly Func<ServiceRequest, Task<ServiceResponse>> _requestHandler;
    private readonly Action<ServiceRequest, ServiceResponse?>? _requestAbandoned;
    private readonly Func<TimeSpan, CancellationToken, Task> _delay;
    private readonly Func<string, NamedPipeServerStream> _serverFactory;
    private readonly TaskCompletionSource _acceptLoopStopped =
        new(TaskCreationOptions.RunContinuationsAsynchronously);
    private int _activeRequests;
    private bool _disposed;

    internal DaemonHost(
        Func<ServiceRequest, Task<ServiceResponse>> requestHandler,
        Func<TimeSpan, CancellationToken, Task>? delay = null,
        Func<string, NamedPipeServerStream>? serverFactory = null,
        Action<ServiceRequest, ServiceResponse?>? requestAbandoned = null)
    {
        ArgumentNullException.ThrowIfNull(requestHandler);
        _requestHandler = requestHandler;
        _requestAbandoned = requestAbandoned;
        _delay = delay
            ?? Task.Delay;
        _serverFactory = serverFactory ?? ServiceSecurity.CreateSecureServer;
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "The pipe accept-loop retries listener failures with bounded backoff so the daemon remains available; this boundary never converts a workbook operation failure into success.")]
    internal async Task RunAsync(string pipeName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(pipeName);
        var currentBackoff = InitialBackoff;
        while (!_shutdownCts.IsCancellationRequested)
        {
            NamedPipeServerStream? server = null;
            try
            {
                server = _serverFactory(pipeName);
                await server.WaitForConnectionAsync(_shutdownCts.Token);
                currentBackoff = InitialBackoff;
                if (_activeConnectionTasks.Count >= MaximumConnections)
                {
                    System.Diagnostics.Trace.TraceWarning("Daemon connection capacity reached; refusing a new connection.");
                    continue;
                }

                var clientServer = server;
                server = null;
                var connectionTask = RunConnectionAsync(clientServer);
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

    internal Task WaitForAcceptLoopStoppedAsync() => _acceptLoopStopped.Task;

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "A disconnected or faulted RPC connection is diagnosed and its pipe disposed without taking down unrelated clients. Request errors are transported by StreamJsonRpc before this connection boundary.")]
    private async Task RunConnectionAsync(
        NamedPipeServerStream clientServer)
    {
        try
        {
            var rpcTarget = new DaemonRpcTarget(
                DispatchRequestAsync, _requestAbandoned);
            using var rpc = new JsonRpc(clientServer, clientServer, rpcTarget)
            {
                CancelLocallyInvokedMethodsWhenConnectionIsClosed = true
            };
            rpc.StartListening();
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

    private async Task<ServiceResponse> DispatchRequestAsync(ServiceRequest request)
    {
        var sessionId = request.SessionId;
        var claimSession = !string.IsNullOrWhiteSpace(sessionId)
            && request.Command is not ("window.show" or "window.hide");
        if (claimSession && !_busySessions.TryAdd(sessionId!, 0))
        {
            return BusyResponse(request, "Another request is still running for this session. " +
                "Inspect session list for its state before submitting more work.");
        }
        var control = request.Command is "service.ping" or "service.status" or "session.list"
            or "window.show" or "window.hide";
        var admitted = false;
        try
        {
            if (!control)
            {
                if (Interlocked.Increment(ref _activeRequests) > 10)
                {
                    Interlocked.Decrement(ref _activeRequests);
                    return BusyResponse(request, "The service is at its request capacity. No Excel work was started.");
                }
                admitted = true;
            }
            return await _requestHandler(request);
        }
        finally
        {
            if (admitted)
                Interlocked.Decrement(ref _activeRequests);
            if (claimSession)
                _busySessions.TryRemove(sessionId!, out _);
        }
    }

    private static ServiceResponse BusyResponse(ServiceRequest request, string message) => new()
    {
        Success = false,
        Command = request.Command,
        SessionId = request.SessionId,
        ErrorCategory = "Busy",
        ErrorMessage = message
    };

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
