using System.IO.Pipes;
using System.Text.Json;
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
    private readonly TimeProvider _timeProvider;
    private bool _disposed;
    private readonly bool _useControlEndpoint;

    public static readonly TimeSpan DefaultConnectTimeout = TimeSpan.FromSeconds(5);
    public static readonly TimeSpan DefaultRequestTimeout =
        TimeSpan.FromSeconds(ParameterTransforms.MaximumTimeoutSeconds + 60);

    /// <summary>Whether this client has established a connection to the service.</summary>
    public bool HasConnected { get; private set; }

    public ServiceClient(string pipeName, TimeSpan? connectTimeout = null, TimeSpan? requestTimeout = null,
        bool useControlEndpoint = false)
        : this(
            pipeName,
            connectTimeout,
            requestTimeout,
            TimeProvider.System,
            useControlEndpoint)
    {
    }

    internal ServiceClient(
        string pipeName,
        TimeSpan? connectTimeout,
        TimeSpan? requestTimeout,
        TimeProvider timeProvider,
        bool useControlEndpoint = false)
    {
        ArgumentNullException.ThrowIfNull(timeProvider);
        _pipeName = pipeName;
        _connectTimeout = connectTimeout ?? DefaultConnectTimeout;
        _requestTimeout = requestTimeout ?? DefaultRequestTimeout;
        _timeProvider = timeProvider;
        _useControlEndpoint = useControlEndpoint;
    }

    internal static bool IsControlCommand(string command) =>
        command is "service.ping" or "service.status" or "service.shutdown" or "session.list"
            or "window.show" or "window.hide";

    internal static string GetControlPipeName(string pipeName) => pipeName + "-control";

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

        var startedAt = _timeProvider.GetTimestamp();
        using var pipe = ServiceSecurity.CreateClient(
            _useControlEndpoint && IsControlCommand(request.Command) ? GetControlPipeName(_pipeName) : _pipeName);
        var connected = false;

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            var connectTimeout = GetStepTimeout(
                _connectTimeout,
                totalTimeout,
                startedAt,
                _timeProvider);
            using (var connectTimeoutCts = new CancellationTokenSource(
                connectTimeout,
                _timeProvider))
            using (var connectCts = CancellationTokenSource.CreateLinkedTokenSource(
                cancellationToken,
                connectTimeoutCts.Token))
            {
                await pipe.ConnectAsync(ToTimeoutMilliseconds(connectTimeout), connectCts.Token);
            }

            connected = true;
            HasConnected = true;

            // Use StreamJsonRpc typed proxy for the RPC call
            var proxy = JsonRpc.Attach<IExcelDaemonRpc>(pipe);
            try
            {
                var requestTimeout = GetStepTimeout(
                    _requestTimeout,
                    totalTimeout,
                    startedAt,
                    _timeProvider);
                using var requestTimeoutCts = new CancellationTokenSource(
                    requestTimeout,
                    _timeProvider);
                using var requestCts = CancellationTokenSource.CreateLinkedTokenSource(
                    cancellationToken,
                    requestTimeoutCts.Token);
                using var disconnectMonitorCts = CancellationTokenSource.CreateLinkedTokenSource(requestCts.Token);

                var callTask = proxy.ProcessCommandAsync(request, requestCts.Token);
                var disconnectTask = WaitForPipeDisconnectAsync(
                    pipe,
                    _timeProvider,
                    disconnectMonitorCts.Token);
                var completed = await Task.WhenAny(callTask, disconnectTask);
                if (completed == disconnectTask
                    && disconnectTask.IsCompletedSuccessfully
                    && disconnectTask.Result
                    && !callTask.IsCompleted)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    return CreateConnectionLostResponse(request);
                }

                await disconnectMonitorCts.CancelAsync();
                var response = await callTask.WaitAsync(requestCts.Token);
                if (cancellationToken.IsCancellationRequested)
                {
                    await RevealCompletedSessionAsync(request, response);
                    cancellationToken.ThrowIfCancellationRequested();
                }
                return response;
            }
            finally
            {
                // Dispose the underlying JsonRpc to clean up the connection
                ((IDisposable)proxy).Dispose();
            }
        }
        catch (TimeoutException)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return CreateTimeoutResponse(request, connected, nameof(TimeoutException));
        }
        catch (OperationCanceledException ex) when (cancellationToken.IsCancellationRequested)
        {
            throw new OperationCanceledException(ex.Message, ex, cancellationToken);
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            return CreateTimeoutResponse(request, connected, nameof(OperationCanceledException));
        }
        catch (ConnectionLostException)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return CreateConnectionLostResponse(request);
        }
        catch (DirectoryNotFoundException ex) when (!connected)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return new ServiceResponse
            {
                Success = false,
                Command = request.Command,
                SessionId = request.SessionId,
                ErrorCategory = "ServiceUnavailable",
                ErrorMessage = $"Cannot connect to service: {ex.Message}",
                ExceptionType = nameof(DirectoryNotFoundException)
            };
        }
        catch (IOException ex) when (ex.Message.Contains("pipe"))
        {
            cancellationToken.ThrowIfCancellationRequested();
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

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "A completed-session recovery failure is diagnosed without replacing caller cancellation or saving, closing, or replaying workbook work.")]
    private async Task RevealCompletedSessionAsync(ServiceRequest request, ServiceResponse response)
    {
        if (!response.Success || request.Command is not ("session.open" or "session.create"))
            return;

        try
        {
            var sessionId = response.SessionId;
            if (string.IsNullOrWhiteSpace(sessionId))
            {
                using var result = JsonDocument.Parse(response.Result
                    ?? throw new InvalidOperationException("Session creation returned no result."));
                sessionId = result.RootElement.GetProperty("sessionId").GetString();
            }
            if (string.IsNullOrWhiteSpace(sessionId))
                throw new InvalidOperationException("Session creation returned no session ID.");

            // The daemon has finished observing this completed request. Reveal its
            // retained workbook before cancellation discards the reply/session ID.
            using var recoveryClient = new ServiceClient(_pipeName, useControlEndpoint: _useControlEndpoint);
            var recovery = await recoveryClient.SendAsync(new ServiceRequest
            {
                Command = "window.show",
                SessionId = sessionId
            }, TimeSpan.FromSeconds(5), CancellationToken.None);
            if (!recovery.Success)
                System.Diagnostics.Trace.TraceError("Could not reveal completed Excel session after cancellation: {0}", recovery.ErrorMessage);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Trace.TraceError("Could not reveal completed Excel session after cancellation: {0}", ex);
        }
    }

    internal static TimeSpan GetStepTimeout(
        TimeSpan configuredTimeout,
        TimeSpan? totalTimeout,
        long startedAt,
        TimeProvider timeProvider)
    {
        if (totalTimeout is null)
        {
            return configuredTimeout;
        }

        var remaining = totalTimeout.Value - timeProvider.GetElapsedTime(startedAt);
        if (remaining <= TimeSpan.Zero)
        {
            throw new TimeoutException();
        }

        return configuredTimeout != Timeout.InfiniteTimeSpan && configuredTimeout <= remaining
            ? configuredTimeout
            : remaining;
    }

    internal static int ToTimeoutMilliseconds(TimeSpan timeout) =>
        timeout == Timeout.InfiniteTimeSpan ? Timeout.Infinite
            : Math.Max(1, checked((int)Math.Ceiling(timeout.TotalMilliseconds)));

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

    private static async Task<bool> WaitForPipeDisconnectAsync(
        Stream pipe,
        TimeProvider timeProvider,
        CancellationToken cancellationToken)
    {
        try
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                if (pipe is PipeStream pipeStream && !pipeStream.IsConnected)
                {
                    return true;
                }

                await Task.Delay(
                    TimeSpan.FromMilliseconds(250),
                    timeProvider,
                    cancellationToken);
            }
        }
        catch (OperationCanceledException)
        {
            return false; // Caller completion or timeout is not a disconnect signal.
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
