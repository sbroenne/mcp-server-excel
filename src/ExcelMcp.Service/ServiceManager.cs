namespace Sbroenne.ExcelMcp.Service;

/// <summary>
/// Manages ExcelMCP Service lifecycle for in-process hosting.
/// The service runs within the host process (MCP Server or CLI),
/// communicating via named pipe for cross-thread STA marshalling.
/// </summary>
public static class ServiceManager
{
    private static ExcelMcpService? _service;
    private static Task? _serviceTask;
    private static readonly SemaphoreSlim _startLock = new(1, 1);

    /// <summary>
    /// Ensures the in-process service is running, starting it if necessary.
    /// Thread-safe — only one service instance will be created.
    /// </summary>
    public static async Task<bool> EnsureServiceRunningAsync(CancellationToken cancellationToken = default)
    {
        // Fast path: already running
        if (_service != null && _serviceTask != null && !_serviceTask.IsCompleted)
        {
            return true;
        }

        await _startLock.WaitAsync(cancellationToken);
        try
        {
            // Double-check after acquiring lock
            if (_service != null && _serviceTask != null && !_serviceTask.IsCompleted)
            {
                return true;
            }

            // Clean up any previous failed instance
            if (_service != null)
            {
                _service.Dispose();
                _service = null;
                _serviceTask = null;
            }

            _service = new ExcelMcpService();
            _serviceTask = Task.Run(() => _service.RunAsync(), cancellationToken);

            // Wait for the pipe server to be ready
            for (int i = 0; i < 20; i++)
            {
                await Task.Delay(100, cancellationToken);
                using var client = new ServiceClient(connectTimeout: TimeSpan.FromSeconds(1));
                if (await client.PingAsync(cancellationToken))
                {
                    return true;
                }
            }

            return false;
        }
        catch (Exception)
        {
            return false;
        }
        finally
        {
            _startLock.Release();
        }
    }

    /// <summary>
    /// Stops the in-process service.
    /// </summary>
    public static async Task<bool> StopServiceAsync(CancellationToken cancellationToken = default)
    {
        if (_service == null)
        {
            return true;
        }

        _service.RequestShutdown();

        if (_serviceTask != null)
        {
            try
            {
                await _serviceTask.WaitAsync(TimeSpan.FromSeconds(5), cancellationToken);
            }
            catch (TimeoutException)
            {
                // Service didn't stop in time — dispose anyway
            }
        }

        _service.Dispose();
        _service = null;
        _serviceTask = null;

        return true;
    }

    /// <summary>
    /// Gets service status information.
    /// </summary>
    public static Task<ServiceStatus> GetStatusAsync(CancellationToken cancellationToken = default)
    {
        _ = cancellationToken; // Reserved for future use
        if (_service == null || _serviceTask == null || _serviceTask.IsCompleted)
        {
            return Task.FromResult(new ServiceStatus { Running = false });
        }

        return Task.FromResult(new ServiceStatus
        {
            Running = true,
            ProcessId = Environment.ProcessId,
            SessionCount = _service.SessionCount,
            StartTime = _service.StartTime
        });
    }
}

/// <summary>
/// Service status information.
/// </summary>
public sealed class ServiceStatus
{
    public bool Running { get; init; }
    public int ProcessId { get; init; }
    public int SessionCount { get; init; }
    public DateTime StartTime { get; init; }
    public TimeSpan Uptime => Running ? DateTime.UtcNow - StartTime : TimeSpan.Zero;
}
