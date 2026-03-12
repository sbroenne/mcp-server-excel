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

/// <summary>
/// Bridge that holds the in-process ExcelMCP Service for direct method calls.
/// No named pipe — MCP tools call the service directly (same process).
/// </summary>
public static class ServiceBridge
{
    private static readonly SemaphoreSlim _initLock = new(1, 1);
    private static readonly Func<IServiceBridgeBackend> DefaultServiceFactory =
        static () => new ExcelMcpServiceBackend(new Service.ExcelMcpService());

    private static IServiceBridgeBackend? _service;
    private static Func<IServiceBridgeBackend> _serviceFactory = DefaultServiceFactory;

    /// <summary>
    /// JSON serializer options for deserializing service responses.
    /// </summary>
    public static readonly JsonSerializerOptions JsonOptions = ServiceProtocol.JsonOptions;

    /// <summary>
    /// Ensures the in-process ExcelMCP Service is created.
    /// Called automatically on first request.
    /// </summary>
    public static async Task<bool> EnsureServiceAsync(CancellationToken cancellationToken = default)
    {
        if (_service != null)
        {
            return true;
        }

        await _initLock.WaitAsync(cancellationToken);
        try
        {
            if (_service != null)
            {
                return true;
            }

            _service = _serviceFactory();
            return true;
        }
        catch (Exception)
        {
            return false;
        }
        finally
        {
            _initLock.Release();
        }
    }

    /// <summary>
    /// Sends a command to the ExcelMCP Service directly (in-process, no pipe).
    /// </summary>
    public static async Task<ServiceResponse> SendAsync(
        string command,
        string? sessionId = null,
        object? args = null,
        int? timeoutSeconds = null,
        CancellationToken cancellationToken = default)
    {
        if (!await EnsureServiceAsync(cancellationToken))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorMessage = "Failed to start ExcelMCP Service in-process."
            };
        }

        var request = new ServiceRequest
        {
            Command = command,
            SessionId = sessionId,
            Args = args != null ? JsonSerializer.Serialize(args, JsonOptions) : null
        };

        var service = _service!;
        var processTask = Task.Run(async () => await service.ProcessAsync(request), CancellationToken.None);

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
            CleanupCancelledRequest(service, sessionId);

            if (timeoutSeconds.HasValue && !cancellationToken.IsCancellationRequested)
            {
                return new ServiceResponse
                {
                    Success = false,
                    ErrorMessage = $"Operation timed out after {timeoutSeconds} seconds."
                };
            }

            return new ServiceResponse
            {
                Success = false,
                ErrorMessage = string.IsNullOrWhiteSpace(sessionId)
                    ? "Operation was cancelled. The Excel MCP service was reset to avoid leaving a stuck Excel operation behind."
                    : "Operation was cancelled and the session has been closed to avoid leaving a stuck Excel operation behind. Please reopen the file with a new session."
            };
        }
    }

    private static void CleanupCancelledRequest(IServiceBridgeBackend service, string? sessionId)
    {
        if (!string.IsNullOrWhiteSpace(sessionId))
        {
            try
            {
                if (service.ForceCloseSession(sessionId))
                {
                    return;
                }
            }
            catch (Exception)
            {
                // Fall back to resetting the entire service below.
            }
        }

        Dispose();
    }

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
        bool macroEnabled = false,
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
        bool save = true,
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
    /// Saves a session via the service.
    /// </summary>
    public static async Task<ServiceResponse> SaveSessionAsync(
        string sessionId,
        CancellationToken cancellationToken = default)
    {
        return await SendAsync("session.save", sessionId, cancellationToken: cancellationToken);
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
        var service = Interlocked.Exchange(ref _service, null);
        service?.Dispose();
    }

    internal static void SetServiceFactoryForTests(Func<IServiceBridgeBackend> serviceFactory)
    {
        Dispose();
        _serviceFactory = serviceFactory;
    }

    internal static void ResetForTests()
    {
        Dispose();
        _serviceFactory = DefaultServiceFactory;
    }
}
