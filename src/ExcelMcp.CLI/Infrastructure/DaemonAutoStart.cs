using System.Diagnostics;
using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>
/// Ensures the CLI daemon is running before sending commands.
/// Auto-starts the daemon if not already running.
/// </summary>
internal static class DaemonAutoStart
{
    /// <summary>
    /// Gets the pipe name for the CLI daemon (supports env var override for testing).
    /// </summary>
    public static string GetPipeName() =>
        Environment.GetEnvironmentVariable("EXCELMCP_CLI_PIPE") ?? ServiceSecurity.GetCliPipeName();

    /// <summary>
    /// Ensures the daemon is running and returns a connected ServiceClient.
    /// If the daemon is not running, starts it and waits for it to be ready.
    /// </summary>
    public static async Task<ServiceClient> EnsureAndConnectAsync(CancellationToken cancellationToken = default)
    {
        var pipeName = GetPipeName();

        // Try connecting first — fast path when daemon is already running
        var client = new ServiceClient(pipeName, connectTimeout: TimeSpan.FromSeconds(2));
        if (await client.PingAsync(cancellationToken))
        {
            return client;
        }
        client.Dispose();

        // Daemon not running — start it
        await StartDaemonAsync(pipeName, cancellationToken);

        // Return new client connected to the now-running daemon
        return new ServiceClient(pipeName);
    }

    /// <summary>
    /// Starts the CLI daemon process and waits for it to become ready.
    /// </summary>
    private static async Task StartDaemonAsync(string pipeName, CancellationToken cancellationToken)
    {
        var exePath = Environment.ProcessPath;
        if (string.IsNullOrEmpty(exePath))
        {
            throw new InvalidOperationException("Cannot determine executable path to start daemon.");
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = exePath,
            Arguments = "service run",
            UseShellExecute = true,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden
        };

        try
        {
            Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to start daemon: {ex.Message}", ex);
        }

        // Wait for daemon to be ready (up to 5 seconds)
        for (int i = 0; i < 20; i++)
        {
            await Task.Delay(250, cancellationToken);
            using var checkClient = new ServiceClient(pipeName, connectTimeout: TimeSpan.FromSeconds(1));
            if (await checkClient.PingAsync(cancellationToken))
            {
                return;
            }
        }

        throw new TimeoutException("Daemon started but not responding within 5 seconds.");
    }
}
