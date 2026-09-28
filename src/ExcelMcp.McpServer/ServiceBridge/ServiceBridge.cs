using System.Text.Json;
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

internal sealed class ServiceBridgeLifetime : IDisposable
{
    private readonly SemaphoreSlim _initLock = new(1, 1);
    private Func<IServiceBridgeBackend> _serviceFactory;
    private ServiceInstance? _current;
    private Exception? _lastStartupException;
    private long _pendingOwnerToken;
    private long _configurationGeneration;

    internal ServiceBridgeLifetime(Func<IServiceBridgeBackend> serviceFactory)
    {
        _serviceFactory = serviceFactory;
    }

    internal async Task<bool> EnsureServiceAsync(CancellationToken cancellationToken)
    {
        return await AcquireServiceAsync(cancellationToken) != null;
    }

    private async Task<ServiceInstance?> AcquireServiceAsync(
        CancellationToken cancellationToken)
    {
        var current = Volatile.Read(ref _current);
        if (current != null)
        {
            return current;
        }

        await _initLock.WaitAsync(cancellationToken);
        try
        {
            while (true)
            {
                current = Volatile.Read(ref _current);
                if (current != null)
                {
                    return current;
                }

                var generation = Interlocked.Read(ref _configurationGeneration);
                var serviceFactory = Volatile.Read(ref _serviceFactory);
                IServiceBridgeBackend backend;
                try
                {
                    backend = serviceFactory();
                }
                catch (Exception ex)
                {
                    if (generation != Interlocked.Read(ref _configurationGeneration))
                    {
                        continue;
                    }

                    _lastStartupException = ex;
                    return null;
                }

                if (generation != Interlocked.Read(ref _configurationGeneration))
                {
                    backend.Dispose();
                    continue;
                }

                current = new ServiceInstance(
                    backend,
                    Interlocked.Read(ref _pendingOwnerToken));
                Volatile.Write(ref _current, current);
                _lastStartupException = null;
                return current;
            }
        }
        finally
        {
            _initLock.Release();
        }
    }

    internal async Task<ServiceResponse> SendAsync(
        string command,
        string? sessionId,
        object? args,
        int? timeoutSeconds,
        CancellationToken cancellationToken)
    {
        ServiceInstance? instance;
        do
        {
            instance = await AcquireServiceAsync(cancellationToken);
        }
        while (instance != null && !instance.TryAcquire());

        if (instance == null)
        {
            return new ServiceResponse
            {
                Success = false,
                Command = command,
                SessionId = sessionId,
                ErrorCategory = "ServiceStartup",
                ErrorMessage = BuildServiceStartupErrorMessage(_lastStartupException),
                ExceptionType = _lastStartupException?.GetType().Name
            };
        }

        try
        {
            var request = new ServiceRequest
            {
                Command = command,
                SessionId = sessionId,
                Args = args != null ? JsonSerializer.Serialize(args, ServiceBridge.JsonOptions) : null
            };

            var processTask = Task.Run(
                async () => await instance.Backend.ProcessAsync(request),
                CancellationToken.None);

            if (!timeoutSeconds.HasValue && !cancellationToken.CanBeCanceled)
            {
                return await processTask;
            }

            using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            if (timeoutSeconds.HasValue)
            {
                cts.CancelAfter(TimeSpan.FromSeconds(timeoutSeconds.Value));
            }

            try
            {
                return await processTask.WaitAsync(cts.Token);
            }
            catch (OperationCanceledException) when (cts.IsCancellationRequested)
            {
                var completedResponse = await TryGetCompletedResponseAsync(processTask);
                if (completedResponse != null)
                {
                    return completedResponse;
                }

                CleanupCancelledRequest(instance, sessionId);

                if (timeoutSeconds.HasValue && !cancellationToken.IsCancellationRequested)
                {
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

                return new ServiceResponse
                {
                    Success = false,
                    Command = command,
                    SessionId = sessionId,
                    ErrorCategory = "Cancelled",
                    ErrorMessage = string.IsNullOrWhiteSpace(sessionId)
                        ? "Operation was cancelled. The Excel MCP service was reset to avoid leaving a stuck Excel operation behind."
                        : "Operation was cancelled and the session has been closed to avoid leaving a stuck Excel operation behind. Please reopen the file with a new session.",
                    ExceptionType = nameof(OperationCanceledException)
                };
            }
        }
        finally
        {
            instance.Release();
        }
    }

    internal void SetOwnerToken(long ownerToken)
    {
        Interlocked.Exchange(ref _pendingOwnerToken, ownerToken);
    }

    internal bool DisposeIfOwnedBy(long ownerToken)
    {
        if (ownerToken == 0)
        {
            return false;
        }

        var instance = Volatile.Read(ref _current);
        if (instance == null || instance.OwnerToken != ownerToken)
        {
            return false;
        }

        return DisposeIfCurrent(instance);
    }

    internal void SetServiceFactory(Func<IServiceBridgeBackend> serviceFactory)
    {
        ArgumentNullException.ThrowIfNull(serviceFactory);
        Volatile.Write(ref _serviceFactory, serviceFactory);
        Dispose();
    }

    public void Dispose()
    {
        Interlocked.Increment(ref _configurationGeneration);
        var instance = Interlocked.Exchange(ref _current, null);
        instance?.RequestDispose();
        _lastStartupException = null;
    }

    private void CleanupCancelledRequest(ServiceInstance instance, string? sessionId)
    {
        if (!string.IsNullOrWhiteSpace(sessionId))
        {
            try
            {
                if (instance.Backend.ForceCloseSession(sessionId))
                {
                    return;
                }
            }
            catch (Exception)
            {
                // Fall back to resetting this backend generation below.
            }
        }

        DisposeIfCurrent(instance);
    }

    private bool DisposeIfCurrent(ServiceInstance instance)
    {
        if (Interlocked.CompareExchange(ref _current, null, instance) != instance)
        {
            return false;
        }

        instance.RequestDispose();
        _lastStartupException = null;
        return true;
    }

    private static async Task<ServiceResponse?> TryGetCompletedResponseAsync(
        Task<ServiceResponse> processTask)
    {
        if (processTask.IsCompleted)
        {
            return await processTask;
        }

        var completedTask = await Task.WhenAny(
            processTask,
            Task.Delay(TimeSpan.FromMilliseconds(50))).ConfigureAwait(false);

        return completedTask == processTask
            ? await processTask.ConfigureAwait(false)
            : null;
    }

    private static string BuildServiceStartupErrorMessage(Exception? exception)
    {
        if (exception == null)
        {
            return "Failed to start ExcelMCP Service in-process.";
        }

        return $"Failed to start ExcelMCP Service in-process: {exception.GetType().Name}: {exception.Message}";
    }

    private sealed class ServiceInstance(
        IServiceBridgeBackend backend,
        long ownerToken)
    {
        private int _leases;
        private int _disposeRequested;
        private int _disposed;

        internal IServiceBridgeBackend Backend { get; } = backend;
        internal long OwnerToken { get; } = ownerToken;

        internal bool TryAcquire()
        {
            if (Volatile.Read(ref _disposeRequested) != 0)
            {
                return false;
            }

            Interlocked.Increment(ref _leases);
            if (Volatile.Read(ref _disposeRequested) == 0)
            {
                return true;
            }

            Release();
            return false;
        }

        internal void Release()
        {
            Interlocked.Decrement(ref _leases);
        }

        internal void RequestDispose()
        {
            Interlocked.Exchange(ref _disposeRequested, 1);
            if (Interlocked.Exchange(ref _disposed, 1) == 0)
            {
                Backend.Dispose();
            }
        }
    }
}

/// <summary>
/// Bridge that holds the in-process ExcelMCP Service for direct method calls.
/// No named pipe — MCP tools call the service directly (same process).
/// </summary>
public static class ServiceBridge
{
    private static readonly Func<IServiceBridgeBackend> DefaultServiceFactory =
        static () => new ExcelMcpServiceBackend(new Service.ExcelMcpService());
    private static readonly ServiceBridgeLifetime Lifetime = new(DefaultServiceFactory);

    /// <summary>
    /// JSON serializer options for deserializing service responses.
    /// </summary>
    public static readonly JsonSerializerOptions JsonOptions = ServiceProtocol.JsonOptions;

    /// <summary>
    /// Ensures the in-process ExcelMCP Service is created.
    /// Called automatically on first request.
    /// </summary>
    public static Task<bool> EnsureServiceAsync(CancellationToken cancellationToken = default) =>
        Lifetime.EnsureServiceAsync(cancellationToken);

    /// <summary>
    /// Sends a command to the ExcelMCP Service directly (in-process, no pipe).
    /// </summary>
    public static Task<ServiceResponse> SendAsync(
        string command,
        string? sessionId = null,
        object? args = null,
        int? timeoutSeconds = null,
        CancellationToken cancellationToken = default) =>
        Lifetime.SendAsync(command, sessionId, args, timeoutSeconds, cancellationToken);

    /// <summary>
    /// Sends a session-scoped command to the service.
    /// </summary>
    public static async Task<ServiceResponse> WithSessionAsync(
        string sessionId,
        string command,
        object? args = null,
        int? timeoutSeconds = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return new ServiceResponse
            {
                Success = false,
                Command = command,
                ErrorCategory = "InvalidInput",
                ErrorMessage = "sessionId is required. Use file 'open' action to start a session."
            };
        }

        return await SendAsync(command, sessionId, args, timeoutSeconds, cancellationToken);
    }

    /// <summary>
    /// Opens a session via the service.
    /// </summary>
    public static async Task<ServiceResponse> OpenSessionAsync(
        string excelPath,
        bool show = false,
        int? timeoutSeconds = null,
        CancellationToken cancellationToken = default)
    {
        return await SendAsync("session.open", null, new
        {
            filePath = excelPath,
            show,
            timeoutSeconds
        }, timeoutSeconds, cancellationToken);
    }

    /// <summary>
    /// Creates a new file and opens a session via the service.
    /// </summary>
    public static async Task<ServiceResponse> CreateSessionAsync(
        string excelPath,
        bool? macroEnabled = null,
        bool show = false,
        int? timeoutSeconds = null,
        CancellationToken cancellationToken = default)
    {
        return await SendAsync("session.create", null, new
        {
            filePath = excelPath,
            macroEnabled,
            show,
            timeoutSeconds
        }, timeoutSeconds, cancellationToken);
    }

    /// <summary>
    /// Closes a session via the service.
    /// </summary>
    public static async Task<ServiceResponse> CloseSessionAsync(
        string sessionId,
        bool save = false,
        CancellationToken cancellationToken = default)
    {
        return await SendAsync("session.close", sessionId, new { save }, cancellationToken: cancellationToken);
    }

    /// <summary>
    /// Lists active sessions via the service.
    /// </summary>
    public static async Task<ServiceResponse> ListSessionsAsync(CancellationToken cancellationToken = default)
    {
        return await SendAsync("session.list", cancellationToken: cancellationToken);
    }

    /// <summary>
    /// Tests if a file can be opened via the service.
    /// </summary>
    public static async Task<ServiceResponse> TestFileAsync(
        string excelPath,
        CancellationToken cancellationToken = default)
    {
        return await SendAsync("session.test", null, new { filePath = excelPath }, cancellationToken: cancellationToken);
    }

    /// <summary>
    /// Disposes the in-process ExcelMCP Service, auto-saving all sessions before shutdown.
    /// Must be called when the MCP server process exits to prevent silent data loss.
    /// </summary>
    public static void Dispose()
    {
        Lifetime.Dispose();
    }

    internal static void SetTestOwnerToken(long ownerToken)
    {
        Lifetime.SetOwnerToken(ownerToken);
    }

    internal static bool DisposeIfOwnedBy(long ownerToken)
    {
        return Lifetime.DisposeIfOwnedBy(ownerToken);
    }

    internal static void SetServiceFactoryForTests(Func<IServiceBridgeBackend> serviceFactory)
    {
        Lifetime.SetServiceFactory(serviceFactory);
    }

    internal static void ResetForTests()
    {
        Dispose();
        Lifetime.SetOwnerToken(0);
        Lifetime.SetServiceFactory(DefaultServiceFactory);
    }
}
