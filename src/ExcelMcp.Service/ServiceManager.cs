using System.Diagnostics;

namespace Sbroenne.ExcelMcp.Service;

/// <summary>
/// Manages ExcelMCP Service lifecycle: start, stop, status.
/// </summary>
public static class ServiceManager
{
    /// <summary>
    /// Known service binary names that can handle 'service run'.
    /// Both MCP Server and CLI can host the service.
    /// </summary>
    private static readonly string[] ServiceBinaryNames =
    [
        "Sbroenne.ExcelMcp.McpServer.exe",
        "excelcli.exe"
    ];

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
        var startInfo = GetServiceStartInfo();
        if (startInfo == null)
        {
            return false;
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

    /// <summary>
    /// Determines how to start the service process.
    /// Priority: 1) Self-launch (current exe), 2) Co-located binary, 3) dotnet run (dev mode).
    /// </summary>
    private static ProcessStartInfo? GetServiceStartInfo()
    {
        var processPath = Environment.ProcessPath;
        var processName = Path.GetFileName(processPath);

        // 1) Self-launch: current process IS a known service binary (MCP Server or CLI)
        if (!string.IsNullOrEmpty(processPath) &&
            ServiceBinaryNames.Any(name => string.Equals(processName, name, StringComparison.OrdinalIgnoreCase)))
        {
            return CreateStartInfo(processPath);
        }

        // 2) Co-located: find a service binary next to the current executable
        //    Handles test hosts, dotnet.exe, or any non-service process
        var baseDir = AppContext.BaseDirectory;
        foreach (var binaryName in ServiceBinaryNames)
        {
            var candidatePath = Path.Combine(baseDir, binaryName);
            if (File.Exists(candidatePath))
            {
                return CreateStartInfo(candidatePath);
            }
        }

        // 3) Development mode: running via 'dotnet run', use 'dotnet <dll> service run'
        var entryAssembly = System.Reflection.Assembly.GetEntryAssembly();
        if (entryAssembly != null &&
            processPath?.EndsWith("dotnet.exe", StringComparison.OrdinalIgnoreCase) == true)
        {
            return new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = $"\"{entryAssembly.Location}\" service run",
                UseShellExecute = true,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
        }

        return null;
    }

    private static ProcessStartInfo CreateStartInfo(string exePath)
    {
        return new ProcessStartInfo
        {
            FileName = exePath,
            Arguments = "service run",
            UseShellExecute = true,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden
        };
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


