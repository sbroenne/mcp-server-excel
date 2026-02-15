using System.Text;

namespace Sbroenne.ExcelMcp.ComInterop.ServiceClient;

/// <summary>
/// Client for communicating with the ExcelMCP Service via named pipe.
/// Used by both CLI and MCP Server to forward requests to the unified service.
/// </summary>
public sealed class ExcelServiceClient : IDisposable
{
    private readonly TimeSpan _connectTimeout;
    private readonly TimeSpan _requestTimeout;
    private readonly string _source;
    private bool _disposed;

    /// <summary>Default timeout for connecting to the service.</summary>
    public static readonly TimeSpan DefaultConnectTimeout = TimeSpan.FromSeconds(5);

    /// <summary>Default timeout for request completion (5 min for long operations).</summary>
    public static readonly TimeSpan DefaultRequestTimeout = TimeSpan.FromSeconds(300);

    /// <summary>
    /// Creates a new service client.
    /// </summary>
    /// <param name="source">Identifies the client source (e.g., "cli", "mcp-server")</param>
    /// <param name="connectTimeout">Optional connection timeout</param>
    /// <param name="requestTimeout">Optional request timeout</param>
    public ExcelServiceClient(string source = "unknown", TimeSpan? connectTimeout = null, TimeSpan? requestTimeout = null)
    {
        _source = source;
        _connectTimeout = connectTimeout ?? DefaultConnectTimeout;
        _requestTimeout = requestTimeout ?? DefaultRequestTimeout;
    }

    /// <summary>
    /// Sends a request to the service and waits for response.
    /// </summary>
    public async Task<ServiceResponse> SendAsync(ServiceRequest request, CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        using var pipe = ServiceSecurity.CreateClient();
        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeoutCts.CancelAfter(_requestTimeout);

        try
        {
            // Connect to service
            await pipe.ConnectAsync((int)_connectTimeout.TotalMilliseconds, timeoutCts.Token);

            using var reader = new StreamReader(pipe, Encoding.UTF8, leaveOpen: true);
            using var writer = new StreamWriter(pipe, Encoding.UTF8, leaveOpen: true) { AutoFlush = true };

            // Send request
            var requestJson = ServiceProtocol.Serialize(request);
            await writer.WriteLineAsync(requestJson.AsMemory(), timeoutCts.Token);

            // Read response
            var responseJson = await reader.ReadLineAsync(timeoutCts.Token);
            if (string.IsNullOrEmpty(responseJson))
            {
                return new ServiceResponse { Success = false, ErrorMessage = "Empty response from service" };
            }

            return ServiceProtocol.Deserialize<ServiceResponse>(responseJson)
                   ?? new ServiceResponse { Success = false, ErrorMessage = "Invalid response format" };
        }
        catch (TimeoutException)
        {
            return new ServiceResponse { Success = false, ErrorMessage = "Service connection timed out" };
        }
        catch (OperationCanceledException) when (timeoutCts.IsCancellationRequested && !cancellationToken.IsCancellationRequested)
        {
            return new ServiceResponse { Success = false, ErrorMessage = "Service request timed out" };
        }
        catch (IOException ex) when (ex.Message.Contains("pipe"))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "Cannot connect to service. Is it running?" };
        }
    }

    /// <summary>
    /// Sends a command to the service with JSON-serialized arguments.
    /// </summary>
    /// <param name="command">Command in format "category.action" (e.g., "session.open", "range.get-values")</param>
    /// <param name="sessionId">Optional session ID for session-scoped commands</param>
    /// <param name="args">Optional arguments object to serialize</param>
    /// <param name="cancellationToken">Cancellation token</param>
    public async Task<ServiceResponse> SendCommandAsync(
        string command,
        string? sessionId = null,
        object? args = null,
        CancellationToken cancellationToken = default)
    {
        var request = new ServiceRequest
        {
            Command = command,
            SessionId = sessionId,
            Args = args != null ? ServiceProtocol.Serialize(args) : null,
            Source = _source
        };

        return await SendAsync(request, cancellationToken);
    }

    /// <summary>
    /// Pings the service to check if it's alive.
    /// </summary>
    public async Task<bool> PingAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            var response = await SendAsync(new ServiceRequest { Command = "service.ping", Source = _source }, cancellationToken);
            return response.Success;
        }
        catch (Exception)
        {
            // Any communication failure — service is not reachable
            return false;
        }
    }

    /// <inheritdoc />
    public void Dispose()
    {
        _disposed = true;
    }
}


