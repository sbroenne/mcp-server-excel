using Sbroenne.ExcelMcp.Service.Rpc;
using StreamJsonRpc;

namespace Sbroenne.ExcelMcp.Service;

/// <summary>
/// Client for communicating with the ExcelMCP CLI daemon via named pipe + StreamJsonRpc.
/// Each call creates a new pipe connection, makes one RPC call, and disconnects.
/// </summary>
public sealed class ServiceClient : IDisposable
{
    private readonly string _pipeName;
    private readonly TimeSpan _connectTimeout;
    private readonly TimeSpan _requestTimeout;
    private bool _disposed;

    public static readonly TimeSpan DefaultConnectTimeout = TimeSpan.FromSeconds(5);
    public static readonly TimeSpan DefaultRequestTimeout = TimeSpan.FromHours(2); // Long enough that any --timeout value wins before the pipe does

    public ServiceClient(string pipeName, TimeSpan? connectTimeout = null, TimeSpan? requestTimeout = null)
    {
        _pipeName = pipeName;
        _connectTimeout = connectTimeout ?? DefaultConnectTimeout;
        _requestTimeout = requestTimeout ?? DefaultRequestTimeout;
    }

    /// <summary>
    /// Sends a request to the service and waits for response via StreamJsonRpc.
    /// </summary>
    public async Task<ServiceResponse> SendAsync(ServiceRequest request, CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        using var pipe = ServiceSecurity.CreateClient(_pipeName);
        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeoutCts.CancelAfter(_requestTimeout);

        try
        {
            await pipe.ConnectAsync((int)_connectTimeout.TotalMilliseconds, timeoutCts.Token);

            // Use StreamJsonRpc typed proxy for the RPC call
            var proxy = JsonRpc.Attach<IExcelDaemonRpc>(pipe);
            try
            {
                return await proxy.ProcessCommandAsync(request);
            }
            finally
            {
                // Dispose the underlying JsonRpc to clean up the connection
                ((IDisposable)proxy).Dispose();
            }
        }
        catch (TimeoutException)
        {
            return new ServiceResponse { Success = false, ErrorMessage = "Service connection timed out" };
        }
        catch (OperationCanceledException) when (timeoutCts.IsCancellationRequested && !cancellationToken.IsCancellationRequested)
        {
            return new ServiceResponse { Success = false, ErrorMessage = "Service request timed out" };
        }
        catch (ConnectionLostException)
        {
            return new ServiceResponse { Success = false, ErrorMessage = "Connection to service lost. Is it running?" };
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

