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
    /// Ensures the CLI daemon is running and returns a connected ServiceClient.
    /// If the daemon is not running, starts it and waits for it to be ready.
    /// </summary>
    public static async Task<ServiceClient> EnsureAndConnectAsync(CancellationToken cancellationToken = default)
    {
        var pipeName = GetPipeName();

        // Fast path: daemon already running and responsive
        var client = new ServiceClient(pipeName, connectTimeout: TimeSpan.FromSeconds(2));
        if (await client.PingAsync(cancellationToken))
        {
            return client;
        }
        client.Dispose();

        // Ping failed — check OS mutex to distinguish "daemon busy" from "daemon not running".
        // The daemon holds this mutex for its entire lifetime, so its presence means
        // the daemon is running but temporarily unresponsive (e.g., during a heavy refresh).
        // This prevents starting a duplicate daemon (and a duplicate tray icon).
        if (IsDaemonMutexHeld(pipeName))
        {
            // Daemon is running but busy — wait up to 10 seconds for it to become responsive
            for (int i = 0; i < 20; i++)
            {
                await Task.Delay(500, cancellationToken);

                // Re-check mutex: if the daemon exited while we waited, stop waiting
                if (!IsDaemonMutexHeld(pipeName))
                    break;

                using var retryClient = new ServiceClient(pipeName, connectTimeout: TimeSpan.FromSeconds(3));
                if (await retryClient.PingAsync(cancellationToken))
                    return new ServiceClient(pipeName);
            }

            // Daemon mutex gone (exited) or still not responding after 10s — start a new one
        }

        // No daemon running — start it
        await StartDaemonAsync(pipeName, cancellationToken);

        // Return new client connected to the now-running daemon
        return new ServiceClient(pipeName);
    }

    /// <summary>
    /// Checks whether a daemon process currently holds the daemon mutex for the given pipe name.
    /// Returns true if a daemon is running (even if temporarily busy).
    /// </summary>
    private static bool IsDaemonMutexHeld(string pipeName)
    {
        try
        {
            // OpenExisting succeeds if any process has this named mutex open.
            // The daemon opens it with initiallyOwned:true and holds it for its entire lifetime.
            using var mutex = Mutex.OpenExisting(GetDaemonMutexName(pipeName));
            return true;
        }
        catch (WaitHandleCannotBeOpenedException)
        {
            return false; // No process has this mutex — daemon is not running
        }
        catch (Exception)
        {
            return false; // Access denied or other error — assume not running
        }
    }

    /// <summary>
    /// Gets the OS mutex name for the CLI daemon identified by its pipe name.
    /// Used by both the daemon (to acquire) and the client (to detect a running daemon).
    /// </summary>
    internal static string GetDaemonMutexName(string pipeName) =>
        $"ExcelMcpCli_{pipeName}";
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
