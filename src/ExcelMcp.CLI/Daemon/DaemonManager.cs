using System.Diagnostics;

namespace Sbroenne.ExcelMcp.CLI.Daemon;

/// <summary>
/// Manages daemon lifecycle: start, stop, status.
/// </summary>
internal static class DaemonManager
{
    /// <summary>
    /// Ensures daemon is running, starting it if necessary.
    /// </summary>
    public static async Task<bool> EnsureDaemonRunningAsync(CancellationToken cancellationToken = default)
    {
        // Check if already running
        if (await IsDaemonRunningAsync(cancellationToken))
        {
            return true;
        }

        // Start daemon
        return await StartDaemonAsync(cancellationToken);
    }

    /// <summary>
    /// Checks if daemon is running and responsive.
    /// </summary>
    public static async Task<bool> IsDaemonRunningAsync(CancellationToken cancellationToken = default)
    {
        // First check lock file
        if (!DaemonSecurity.IsDaemonProcessRunning())
        {
            return false;
        }

        // Then ping
        using var client = new DaemonClient(connectTimeout: TimeSpan.FromSeconds(2));
        return await client.PingAsync(cancellationToken);
    }

    /// <summary>
    /// Starts the daemon as a background process.
    /// </summary>
    public static async Task<bool> StartDaemonAsync(CancellationToken cancellationToken = default)
    {
        ProcessStartInfo startInfo;

        // Check if running via 'dotnet run' (development mode)
        var entryAssembly = System.Reflection.Assembly.GetEntryAssembly();
        var isDotnetRun = entryAssembly != null &&
            Environment.ProcessPath?.EndsWith("dotnet.exe", StringComparison.OrdinalIgnoreCase) == true;

        if (isDotnetRun)
        {
            // Development mode: use 'dotnet <dll> daemon run'
            var dllPath = entryAssembly!.Location;
            startInfo = new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = $"\"{dllPath}\" daemon run",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
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
                Arguments = "daemon run",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
        }

        try
        {
            var process = Process.Start(startInfo);
            if (process == null)
            {
                return false;
            }

            // Wait a bit for daemon to start
            await Task.Delay(500, cancellationToken);

            // Verify it's running
            for (int i = 0; i < 10; i++)
            {
                if (await IsDaemonRunningAsync(cancellationToken))
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
    /// Stops the daemon.
    /// </summary>
    public static async Task<bool> StopDaemonAsync(CancellationToken cancellationToken = default)
    {
        if (!await IsDaemonRunningAsync(cancellationToken))
        {
            return true; // Already stopped
        }

        using var client = new DaemonClient();
        var response = await client.SendAsync(new DaemonRequest { Command = "daemon.shutdown" }, cancellationToken);
        return response.Success;
    }

    /// <summary>
    /// Gets daemon status information.
    /// </summary>
    public static async Task<DaemonStatus> GetStatusAsync(CancellationToken cancellationToken = default)
    {
        var pid = DaemonSecurity.ReadLockFilePid();
        var isRunning = await IsDaemonRunningAsync(cancellationToken);

        if (!isRunning)
        {
            return new DaemonStatus { Running = false };
        }

        using var client = new DaemonClient();
        var response = await client.SendAsync(new DaemonRequest { Command = "daemon.status" }, cancellationToken);

        if (response.Success && response.Result != null)
        {
            var status = DaemonProtocol.Deserialize<DaemonStatus>(response.Result);
            if (status != null)
            {
                return status;
            }
        }

        return new DaemonStatus { Running = true, ProcessId = pid ?? 0 };
    }
}

/// <summary>
/// Daemon status information.
/// </summary>
internal sealed class DaemonStatus
{
    public bool Running { get; init; }
    public int ProcessId { get; init; }
    public int SessionCount { get; init; }
    public DateTime StartTime { get; init; }
    public TimeSpan Uptime => Running ? DateTime.UtcNow - StartTime : TimeSpan.Zero;
}
