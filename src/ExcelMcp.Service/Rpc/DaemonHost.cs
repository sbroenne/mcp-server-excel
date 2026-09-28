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
    private DateTimeOffset _lastActivityTime;
    private bool _disposed;

    internal DaemonHost(
        Func<ServiceRequest, Task<ServiceResponse>> requestHandler,
        Func<int> sessionCount,
        TimeProvider? timeProvider = null)
    {
        ArgumentNullException.ThrowIfNull(requestHandler);
        ArgumentNullException.ThrowIfNull(sessionCount);
        _requestHandler = requestHandler;
        _sessionCount = sessionCount;
        _timeProvider = timeProvider ?? TimeProvider.System;
        _lastActivityTime = _timeProvider.GetUtcNow();
    }

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
                server = ServiceSecurity.CreateSecureServer(pipeName);
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
            catch (Exception)
            {
                try
                {
                    await Task.Delay(currentBackoff, _timeProvider, _shutdownCts.Token);
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
                    catch (Exception)
                    {
                        // The peer may already have disconnected.
                    }

                    await server.DisposeAsync();
                }
            }
        }

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
                await Task.Delay(TimeSpan.FromSeconds(30), _timeProvider, cancellationToken);
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
        }
    }

    private async Task ShutdownAfterResponseAsync()
    {
        try
        {
            await Task.Delay(TimeSpan.FromMilliseconds(100), _timeProvider, _shutdownCts.Token);
            RequestShutdown();
        }
        catch (OperationCanceledException) when (_shutdownCts.IsCancellationRequested)
        {
        }
    }

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
