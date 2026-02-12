using System.Text;

namespace Sbroenne.ExcelMcp.Service;

/// <summary>
/// Client for communicating with the ExcelMCP Service via named pipe.
/// </summary>
public sealed class ServiceClient : IDisposable
{
    private readonly TimeSpan _connectTimeout;
    private readonly TimeSpan _requestTimeout;
    private bool _disposed;

    public static readonly TimeSpan DefaultConnectTimeout = TimeSpan.FromSeconds(5);
    public static readonly TimeSpan DefaultRequestTimeout = TimeSpan.FromSeconds(300); // 5 min for long operations

    public ServiceClient(TimeSpan? connectTimeout = null, TimeSpan? requestTimeout = null)
    {
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
        catch (IOException ex) when (ex.Message.Contains("pipe"))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "Cannot connect to service. Is it running?" };
        }
    }

    /// <summary>
    /// Pings the service to check if it's alive.
    /// </summary>
    public async Task<bool> PingAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            var response = await SendAsync(new ServiceRequest { Command = "service.ping" }, cancellationToken);
            return response.Success;
        }
        catch (Exception)
        {
            // Any other communication failure — service is not reachable
            return false;
        }
    }

    public void Dispose()
    {
        _disposed = true;
    }
}


