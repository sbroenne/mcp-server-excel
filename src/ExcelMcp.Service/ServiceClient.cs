using System.Diagnostics;
using System.IO.Pipes;
using Sbroenne.ExcelMcp.Core.Utilities;
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
    public static readonly TimeSpan DefaultRequestTimeout =
        TimeSpan.FromSeconds(ParameterTransforms.MaximumTimeoutSeconds + 60);

    public ServiceClient(string pipeName, TimeSpan? connectTimeout = null, TimeSpan? requestTimeout = null)
    {
        _pipeName = pipeName;
        _connectTimeout = connectTimeout ?? DefaultConnectTimeout;
        _requestTimeout = requestTimeout ?? DefaultRequestTimeout;
    }

    /// <summary>
    /// Sends a request to the service and waits for response via StreamJsonRpc.
    /// </summary>
    public Task<ServiceResponse> SendAsync(
        ServiceRequest request,
        CancellationToken cancellationToken = default) =>
        SendCoreAsync(request, totalTimeout: null, cancellationToken);

    /// <summary>
    /// Sends a request using one timeout across both pipe connection and response.
    /// </summary>
    public Task<ServiceResponse> SendAsync(
        ServiceRequest request,
        TimeSpan totalTimeout,
        CancellationToken cancellationToken = default)
    {
        ArgumentOutOfRangeException.ThrowIfLessThanOrEqual(totalTimeout, TimeSpan.Zero);
        return SendCoreAsync(request, totalTimeout, cancellationToken);
    }

    private async Task<ServiceResponse> SendCoreAsync(
        ServiceRequest request,
        TimeSpan? totalTimeout,
        CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        var startedAt = Stopwatch.GetTimestamp();
        using var pipe = ServiceSecurity.CreateClient(_pipeName);
        var connectTimeout = GetStepTimeout(_connectTimeout, totalTimeout, startedAt);
        using var connectCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        connectCts.CancelAfter(connectTimeout);
        var connected = false;

        try
        {
            await pipe.ConnectAsync(ToTimeoutMilliseconds(connectTimeout), connectCts.Token);
            connected = true;

            // Use StreamJsonRpc typed proxy for the RPC call
            var proxy = JsonRpc.Attach<IExcelDaemonRpc>(pipe);
            try
            {
                var requestTimeout = GetStepTimeout(_requestTimeout, totalTimeout, startedAt);
                using var requestCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                using var disconnectMonitorCts = CancellationTokenSource.CreateLinkedTokenSource(requestCts.Token);
                requestCts.CancelAfter(requestTimeout);

                var callTask = proxy.ProcessCommandAsync(request);
                var disconnectTask = WaitForPipeDisconnectAsync(pipe, disconnectMonitorCts.Token);
                var completed = await Task.WhenAny(callTask, disconnectTask);
                if (completed == disconnectTask
                    && disconnectTask.IsCompletedSuccessfully
                    && disconnectTask.Result
                    && !callTask.IsCompleted)
                {
                    return CreateConnectionLostResponse(request);
                }

                await disconnectMonitorCts.CancelAsync();
                return await callTask.WaitAsync(requestCts.Token);
            }
            finally
            {
                // Dispose the underlying JsonRpc to clean up the connection
                ((IDisposable)proxy).Dispose();
            }
        }
        catch (TimeoutException)
        {
            return CreateTimeoutResponse(request, connected, nameof(TimeoutException));
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            return CreateTimeoutResponse(request, connected, nameof(OperationCanceledException));
        }
        catch (ConnectionLostException)
        {
            return CreateConnectionLostResponse(request);
        }
        catch (IOException ex) when (ex.Message.Contains("pipe"))
        {
            return new ServiceResponse
            {
                Success = false,
                Command = request.Command,
                SessionId = request.SessionId,
                ErrorCategory = "ServiceUnavailable",
                ErrorMessage = "Cannot connect to service. Is it running?",
                ExceptionType = nameof(IOException)
            };
        }
    }

    private static TimeSpan GetStepTimeout(
        TimeSpan configuredTimeout,
        TimeSpan? totalTimeout,
        long startedAt)
    {
        if (totalTimeout is null)
        {
            return configuredTimeout;
        }

        var remaining = totalTimeout.Value - Stopwatch.GetElapsedTime(startedAt);
        if (remaining <= TimeSpan.Zero)
        {
            throw new TimeoutException();
        }

        return configuredTimeout <= remaining ? configuredTimeout : remaining;
    }

    private static int ToTimeoutMilliseconds(TimeSpan timeout) =>
        Math.Max(1, checked((int)Math.Ceiling(timeout.TotalMilliseconds)));

    private static ServiceResponse CreateTimeoutResponse(
        ServiceRequest request,
        bool connected,
        string exceptionType)
    {
        return new ServiceResponse
        {
            Success = false,
            Command = request.Command,
            SessionId = request.SessionId,
            ErrorCategory = "Timeout",
            ErrorMessage = connected
                ? "Service request timed out"
                : "Service connection timed out",
            ExceptionType = exceptionType
        };
    }

    private static async Task<bool> WaitForPipeDisconnectAsync(Stream pipe, CancellationToken cancellationToken)
    {
        try
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                if (pipe is PipeStream pipeStream && !pipeStream.IsConnected)
                {
                    return true;
                }

                await Task.Delay(TimeSpan.FromMilliseconds(250), cancellationToken);
            }
        }
        catch (OperationCanceledException)
        {
            // Caller completed or request timed out; this is not a disconnect signal.
        }
        catch (ObjectDisposedException)
        {
            return true;
        }

        return false;
    }

    private static ServiceResponse CreateConnectionLostResponse(ServiceRequest request)
    {
        return new ServiceResponse
        {
            Success = false,
            Command = request.Command,
            SessionId = request.SessionId,
            ErrorCategory = "ServiceUnavailable",
            ErrorMessage = "Connection to service lost while waiting for a response. The daemon may have exited or restarted; run 'excelcli service status' or 'excelcli service stop' and retry.",
            ExceptionType = nameof(ConnectionLostException)
        };
    }

    /// <summary>
    /// Pings the service to check if it's alive.
    /// </summary>
    public async Task<bool> PingAsync(CancellationToken cancellationToken = default)
    {
        var response = await SendAsync(new ServiceRequest { Command = "service.ping" }, cancellationToken);
        return response.Success;
    }

    public void Dispose()
    {
        _disposed = true;
    }
}
