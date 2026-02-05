using System.Diagnostics;

namespace Sbroenne.ExcelMcp.Service;

/// <summary>
/// Manages ExcelMCP Service lifecycle: start, stop, status.
/// </summary>
public static class ServiceManager
{
    /// <summary>
    /// Ensures service is running, starting it if necessary.
    /// </summary>
    public static async Task<bool> EnsureServiceRunningAsync(CancellationToken cancellationToken = default)
    {
        // Check if already running
        if (await IsServiceRunningAsync(cancellationToken))
        {
            return true;
        }

        // Start service
        return await StartServiceAsync(cancellationToken);
    }

    /// <summary>
    /// Checks if service is running and responsive.
    /// </summary>
    public static async Task<bool> IsServiceRunningAsync(CancellationToken cancellationToken = default)
    {
        // First check lock file
        if (!ServiceSecurity.IsServiceProcessRunning())
        {
            return false;
        }

        // Then ping
        using var client = new ServiceClient(connectTimeout: TimeSpan.FromSeconds(2));
        return await client.PingAsync(cancellationToken);
    }

    /// <summary>
    /// Starts the service as a background process.
    /// </summary>
    public static async Task<bool> StartServiceAsync(CancellationToken cancellationToken = default)
    {
        ProcessStartInfo startInfo;

        // Check if running via 'dotnet run' (development mode)
        var entryAssembly = System.Reflection.Assembly.GetEntryAssembly();
        var isDotnetRun = entryAssembly != null &&
            Environment.ProcessPath?.EndsWith("dotnet.exe", StringComparison.OrdinalIgnoreCase) == true;

        // Use UseShellExecute=true to create service in separate process group
        // This prevents parent's console handles from being inherited
        if (isDotnetRun)
        {
            // Development mode: use 'dotnet <dll> service run'
            var dllPath = entryAssembly!.Location;
            startInfo = new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = $"\"{dllPath}\" service run",
                UseShellExecute = true,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
        }
        else
        {
            // Production mode: use the exe directly
            var exePath = Environment.ProcessPath;
            if (string.IsNullOrEmpty(exePath))
            {
                return false;
            }

            startInfo = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = "service run",
                UseShellExecute = true,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
        }

        try
        {
            var process = Process.Start(startInfo);
            if (process == null)
            {
                return false;
            }

            // Wait a bit for service to start
            await Task.Delay(500, cancellationToken);

            // Verify it's running
            for (int i = 0; i < 10; i++)
            {
                if (await IsServiceRunningAsync(cancellationToken))
                {
                    return true;
                }
                await Task.Delay(200, cancellationToken);
            }

            return false;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Stops the service.
    /// </summary>
    public static async Task<bool> StopServiceAsync(CancellationToken cancellationToken = default)
    {
        if (!await IsServiceRunningAsync(cancellationToken))
        {
            return true; // Already stopped
        }

        using var client = new ServiceClient();
        var response = await client.SendAsync(new ServiceRequest { Command = "service.shutdown" }, cancellationToken);
        return response.Success;
    }

    /// <summary>
    /// Gets service status information.
    /// </summary>
    public static async Task<ServiceStatus> GetStatusAsync(CancellationToken cancellationToken = default)
    {
        var pid = ServiceSecurity.ReadLockFilePid();
        var isRunning = await IsServiceRunningAsync(cancellationToken);

        if (!isRunning)
        {
            return new ServiceStatus { Running = false };
        }

        using var client = new ServiceClient();
        var response = await client.SendAsync(new ServiceRequest { Command = "service.status" }, cancellationToken);

        if (response.Success && response.Result != null)
        {
            var status = ServiceProtocol.Deserialize<ServiceStatus>(response.Result);
            if (status != null)
            {
                return status;
            }
        }

        return new ServiceStatus { Running = true, ProcessId = pid ?? 0 };
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


