using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.ServiceClient;
using ServiceManager = Sbroenne.ExcelMcp.Service.ServiceManager;

namespace Sbroenne.ExcelMcp.McpServer.ServiceBridge;

/// <summary>
/// Bridge that forwards MCP Server requests to the ExcelMCP Service.
/// This enables unified session management across CLI and MCP.
/// </summary>
public static class ServiceBridge
{
    private static readonly SemaphoreSlim _initLock = new(1, 1);
    private static bool _serviceStarted;

    /// <summary>
    /// JSON serializer options for deserializing service responses.
    /// </summary>
    public static readonly JsonSerializerOptions JsonOptions = ServiceProtocol.JsonOptions;

    /// <summary>
    /// Ensures the ExcelMCP Service is running.
    /// Called automatically on first request.
    /// </summary>
    public static async Task<bool> EnsureServiceAsync(CancellationToken cancellationToken = default)
    {
        if (_serviceStarted && ExcelServiceClient.IsServiceRunning)
        {
            return true;
        }

        await _initLock.WaitAsync(cancellationToken);
        try
        {
            if (_serviceStarted && ExcelServiceClient.IsServiceRunning)
            {
                return true;
            }

            _serviceStarted = await ServiceManager.EnsureServiceRunningAsync(cancellationToken);
            return _serviceStarted;
        }
        finally
        {
            _initLock.Release();
        }
    }

    /// <summary>
    /// Sends a command to the ExcelMCP Service and returns the response.
    /// </summary>
    /// <param name="command">Command in format "category.action"</param>
    /// <param name="sessionId">Optional session ID</param>
    /// <param name="args">Optional arguments to serialize</param>
    /// <param name="timeoutSeconds">Optional timeout in seconds</param>
    /// <param name="cancellationToken">Cancellation token</param>
    /// <returns>Service response</returns>
    public static async Task<ServiceResponse> SendAsync(
        string command,
        string? sessionId = null,
        object? args = null,
        int? timeoutSeconds = null,
        CancellationToken cancellationToken = default)
    {
        // Ensure service is running
        if (!await EnsureServiceAsync(cancellationToken))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorMessage = "Failed to start ExcelMCP Service. Ensure excelcli.exe is installed and accessible."
            };
        }

        var timeout = timeoutSeconds.HasValue
            ? TimeSpan.FromSeconds(timeoutSeconds.Value)
            : ExcelServiceClient.DefaultRequestTimeout;

        using var client = new ExcelServiceClient("mcp-server", requestTimeout: timeout);
        return await client.SendCommandAsync(command, sessionId, args, cancellationToken);
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
}


