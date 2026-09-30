using System.Text.Json;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.McpServer.ServiceBridge;

internal interface IServiceBridgeBackend : IDisposable
{
    Task<ServiceResponse> ProcessAsync(ServiceRequest request);
    bool ForceCloseSession(string sessionId);
}

internal sealed class ExcelMcpServiceBackend(Service.ExcelMcpService service) : IServiceBridgeBackend
{
    public Task<ServiceResponse> ProcessAsync(ServiceRequest request) => service.ProcessAsync(request);
    public bool ForceCloseSession(string sessionId) => service.SessionManager.CloseSession(sessionId, save: false, force: true);
    public void Dispose() => service.Dispose();
}

/// <summary>Host-owned access to the in-process Excel service.</summary>
public sealed class ServiceBridge : IDisposable
{
    private readonly SemaphoreSlim _initLock = new(1, 1);
    private readonly object _stateLock = new();
    private readonly Func<IServiceBridgeBackend> _serviceFactory;
    private readonly ILogger<ServiceBridge> _logger;
    private ServiceInstance? _current;
    private bool _disposed;

    internal ServiceBridge(Func<IServiceBridgeBackend> serviceFactory, ILogger<ServiceBridge>? logger = null)
    {
        _serviceFactory = serviceFactory;
        _logger = logger ?? NullLogger<ServiceBridge>.Instance;
    }

    private async Task<ServiceInstance> AcquireServiceAsync(CancellationToken cancellationToken)
    {
        lock (_stateLock)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (_current is not null)
                return _current;
        }

        await _initLock.WaitAsync(cancellationToken);
        try
        {
            lock (_stateLock)
            {
                ObjectDisposedException.ThrowIf(_disposed, this);
                if (_current is not null)
                    return _current;
            }

            var backend = _serviceFactory();
            lock (_stateLock)
            {
                if (!_disposed)
                {
                    _current = new ServiceInstance(backend);
                    return _current;
                }
            }

            backend.Dispose();
            throw new ObjectDisposedException(nameof(ServiceBridge));
        }
        finally
        {
            _initLock.Release();
        }
    }

    public async Task<ServiceResponse> SendAsync(
        string command,
        string? sessionId = null,
        object? args = null,
        int? timeoutSeconds = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ServiceInstance instance;
        try
        {
            instance = await AcquireServiceAsync(cancellationToken);
        }
        catch (Exception ex) when (ex is not OperationCanceledException and not ObjectDisposedException)
        {
            _logger.LogError(ex, "Failed to start the in-process Excel service.");
            return new ServiceResponse
            {
                Success = false,
                Command = command,
                SessionId = sessionId,
                ErrorCategory = "ServiceStartup",
                ErrorMessage = "Failed to start ExcelMCP Service in-process. See server logs for details.",
                ExceptionType = ex.GetType().Name
            };
        }

        var request = new ServiceRequest
        {
            Command = command,
            SessionId = sessionId,
            Args = args is null ? null : JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions)
        };

        // Service dispatch includes synchronous COM waits. Keep those off the SDK's request loop.
        var processTask = Task.Run(() => instance.Backend.ProcessAsync(request), CancellationToken.None);
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        if (timeoutSeconds.HasValue)
            deadline.CancelAfter(TimeSpan.FromSeconds(timeoutSeconds.Value));

        try
        {
            var response = await processTask.WaitAsync(deadline.Token);
            cancellationToken.ThrowIfCancellationRequested();
            return response;
        }
        catch (OperationCanceledException) when (deadline.IsCancellationRequested)
        {
            if (!cancellationToken.IsCancellationRequested && processTask.IsCompleted)
                return await processTask;

            if (!string.IsNullOrWhiteSpace(sessionId))
                CloseCancelledSession(instance, sessionId);

            // Open/create has no session ID until Excel finishes. Reclaim only its eventual session,
            // not the whole service and unrelated workbooks. Also observe all late task failures.
            _ = ObserveCancelledRequestAsync(instance, processTask, command);
            cancellationToken.ThrowIfCancellationRequested();
            return new ServiceResponse
            {
                Success = false,
                Command = command,
                SessionId = sessionId,
                ErrorCategory = "Timeout",
                ErrorMessage = $"Operation timed out after {timeoutSeconds} seconds.",
                ExceptionType = nameof(TimeoutException)
            };
        }
    }

    private async Task ObserveCancelledRequestAsync(
        ServiceInstance instance, Task<ServiceResponse> task, string command)
    {
        try
        {
            var response = await task;
            if (command is "session.open" or "session.create" && response.Success)
            {
                using var result = JsonDocument.Parse(response.Result
                    ?? throw new InvalidOperationException("Session creation returned no result."));
                var sessionId = result.RootElement.GetProperty("sessionId").GetString();
                if (string.IsNullOrWhiteSpace(sessionId))
                    throw new InvalidOperationException("Session creation returned no session ID.");
                CloseCancelledSession(instance, sessionId);
            }
            else if (!response.Success)
            {
                _logger.LogWarning("Cancelled {Command} finished with error category {ErrorCategory}.",
                    command, response.ErrorCategory);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Cleanup of cancelled {Command} failed.", command);
        }
    }

    private void CloseCancelledSession(ServiceInstance instance, string sessionId)
    {
        try
        {
            if (instance.Backend.ForceCloseSession(sessionId))
                return;
            _logger.LogWarning("Forced session cleanup failed; retiring the affected service instance.");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Forced session cleanup failed; retiring the affected service instance.");
        }

        lock (_stateLock)
        {
            if (ReferenceEquals(_current, instance))
                _current = null;
        }
        instance.Dispose();
    }

    public void Dispose()
    {
        ServiceInstance? instance;
        lock (_stateLock)
        {
            if (_disposed)
                return;
            _disposed = true;
            instance = _current;
            _current = null;
        }
        instance?.Dispose();
    }

    private sealed class ServiceInstance(IServiceBridgeBackend backend) : IDisposable
    {
        private int _disposed;
        internal IServiceBridgeBackend Backend { get; } = backend;

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) == 0)
                Backend.Dispose();
        }
    }
}
