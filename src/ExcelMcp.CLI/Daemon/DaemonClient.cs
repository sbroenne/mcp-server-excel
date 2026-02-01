using System.Text;

namespace Sbroenne.ExcelMcp.CLI.Daemon;

/// <summary>
/// Client for communicating with the Excel daemon via named pipe.
/// </summary>
internal sealed class DaemonClient : IDisposable
{
    private readonly TimeSpan _connectTimeout;
    private readonly TimeSpan _requestTimeout;
    private bool _disposed;

    public static readonly TimeSpan DefaultConnectTimeout = TimeSpan.FromSeconds(5);
    public static readonly TimeSpan DefaultRequestTimeout = TimeSpan.FromSeconds(300); // 5 min for long operations

    public DaemonClient(TimeSpan? connectTimeout = null, TimeSpan? requestTimeout = null)
    {
        _connectTimeout = connectTimeout ?? DefaultConnectTimeout;
        _requestTimeout = requestTimeout ?? DefaultRequestTimeout;
    }

    /// <summary>
    /// Sends a request to the daemon and waits for response.
    /// </summary>
    public async Task<DaemonResponse> SendAsync(DaemonRequest request, CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        using var pipe = DaemonSecurity.CreateClient();
        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeoutCts.CancelAfter(_requestTimeout);

        try
        {
            // Connect to daemon
            await pipe.ConnectAsync((int)_connectTimeout.TotalMilliseconds, timeoutCts.Token);

            using var reader = new StreamReader(pipe, Encoding.UTF8, leaveOpen: true);
            using var writer = new StreamWriter(pipe, Encoding.UTF8, leaveOpen: true) { AutoFlush = true };

            // Send request
            var requestJson = DaemonProtocol.Serialize(request);
            await writer.WriteLineAsync(requestJson.AsMemory(), timeoutCts.Token);

            // Read response
            var responseJson = await reader.ReadLineAsync(timeoutCts.Token);
            if (string.IsNullOrEmpty(responseJson))
            {
                return new DaemonResponse { Success = false, ErrorMessage = "Empty response from daemon" };
            }

            return DaemonProtocol.Deserialize<DaemonResponse>(responseJson)
                   ?? new DaemonResponse { Success = false, ErrorMessage = "Invalid response format" };
        }
        catch (TimeoutException)
        {
            return new DaemonResponse { Success = false, ErrorMessage = "Daemon connection timed out" };
        }
        catch (IOException ex) when (ex.Message.Contains("pipe"))
        {
            return new DaemonResponse { Success = false, ErrorMessage = "Cannot connect to daemon. Is it running?" };
        }
    }

    /// <summary>
    /// Pings the daemon to check if it's alive.
    /// </summary>
    public async Task<bool> PingAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            var response = await SendAsync(new DaemonRequest { Command = "daemon.ping" }, cancellationToken);
            return response.Success;
        }
        catch
        {
            return false;
        }
    }

    public void Dispose()
    {
        _disposed = true;
    }
}
