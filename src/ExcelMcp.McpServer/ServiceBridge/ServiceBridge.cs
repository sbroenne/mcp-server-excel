using System.Text.Json;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.McpServer.ServiceBridge;

internal interface IServiceBridgeBackend : IDisposable
{
    Task<ServiceResponse> ProcessAsync(ServiceRequest request);
    void RevealSession(string sessionId);
}

internal sealed class ExcelMcpServiceBackend(Service.ExcelMcpService service) : IServiceBridgeBackend
{
    public Task<ServiceResponse> ProcessAsync(ServiceRequest request) => service.ProcessAsync(request);
    public void RevealSession(string sessionId) => service.SessionManager.RevealForRecovery(sessionId, waitExpired: true);
    public void Dispose() => service.Dispose();
}

/// <summary>Host-owned access to the in-process Excel service.</summary>
public sealed class ServiceBridge : IDisposable
{
    private readonly SemaphoreSlim _initLock = new(1, 1);
    private readonly object _stateLock = new();
    private readonly object _disposeLock = new();
    private readonly Func<IServiceBridgeBackend> _serviceFactory;
    private readonly ILogger<ServiceBridge> _logger;
    private IServiceBridgeBackend? _current;
    private bool _disposed;

    internal ServiceBridge(Func<IServiceBridgeBackend> serviceFactory, ILogger<ServiceBridge>? logger = null)
    {
        _serviceFactory = serviceFactory;
        _logger = logger ?? NullLogger<ServiceBridge>.Instance;
    }

    private async Task<IServiceBridgeBackend> AcquireServiceAsync(CancellationToken cancellationToken)
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
                _current = backend;
                if (!_disposed)
                {
                    return _current;
                }
            }

            Dispose();
            throw new ObjectDisposedException(nameof(ServiceBridge));
        }
        finally
        {
            _initLock.Release();
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Service startup failures are logged and transported as Success=false ServiceStartup responses; caller cancellation and disposal are explicitly excluded from this boundary.")]
    public async Task<ServiceResponse> SendAsync(
        string command,
        string? sessionId = null,
        object? args = null,
        int? timeoutSeconds = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        IServiceBridgeBackend instance;
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
        var processTask = Task.Run(() => instance.ProcessAsync(request), CancellationToken.None);
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        if (timeoutSeconds.HasValue && command is not ("session.open" or "session.create" or "session.test"))
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
                RevealSession(instance, sessionId);

            _ = ObserveCancelledRequestAsync(instance, processTask, command);
            cancellationToken.ThrowIfCancellationRequested();
            return new ServiceResponse
            {
                Success = false,
                Command = command,
                SessionId = sessionId,
                ErrorCategory = "Timeout",
                ErrorMessage = $"The wait expired after {timeoutSeconds} seconds. Excel work may still be running; no session was closed. Inspect session status for its outcome.",
                ExceptionType = nameof(TimeoutException)
            };
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Observes late completion after the caller stopped waiting; errors are logged and a late opened session is revealed, never closed.")]
    private async Task ObserveCancelledRequestAsync(
        IServiceBridgeBackend instance, Task<ServiceResponse> task, string command)
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
                RevealSession(instance, sessionId);
            }
            else if (!response.Success)
            {
                if (response.SessionId is { } pendingSession)
                    RevealSession(instance, pendingSession);
                _logger.LogWarning("Cancelled {Command} finished with error category {ErrorCategory}.",
                    command, response.ErrorCategory);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Late observation of cancelled {Command} failed; workbook ownership is retained.", command);
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Failure to reveal Excel is diagnosed without replacing the caller's cancellation or disposing the Service that owns pending work.")]
    private void RevealSession(IServiceBridgeBackend instance, string sessionId)
    {
        try
        {
            instance.RevealSession(sessionId);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Could not reveal session {SessionId}; workbook ownership is retained.", sessionId);
        }
    }

    public void Dispose()
    {
        lock (_disposeLock)
        {
            IServiceBridgeBackend? instance;
            lock (_stateLock)
            {
                _disposed = true;
                instance = _current;
            }
            instance?.Dispose();
            lock (_stateLock)
            {
                if (ReferenceEquals(_current, instance))
                    _current = null;
            }
        }
    }
}
