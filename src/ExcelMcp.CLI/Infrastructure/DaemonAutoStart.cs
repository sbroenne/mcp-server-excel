using System.Diagnostics;
using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>
/// Ensures the CLI daemon is running before sending commands.
/// Auto-starts the daemon if not already running.
/// </summary>
internal static class DaemonAutoStart
{
    internal static readonly TimeSpan InitialPingTimeout = TimeSpan.FromSeconds(2);
    internal static readonly TimeSpan BusyDaemonConnectTimeout = TimeSpan.FromSeconds(3);
    internal static readonly TimeSpan BusyDaemonRetryInterval = TimeSpan.FromMilliseconds(500);
    internal static readonly TimeSpan BusyDaemonWaitTimeout = TimeSpan.FromSeconds(10);
    internal static readonly TimeSpan StartupReadyConnectTimeout = TimeSpan.FromSeconds(1);
    internal static readonly TimeSpan StartupReadyRetryInterval = TimeSpan.FromMilliseconds(250);
    internal static readonly TimeSpan StartupReadyTimeout = TimeSpan.FromSeconds(10);

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
        if (await PingAsync(pipeName, InitialPingTimeout, cancellationToken))
        {
            return new ServiceClient(pipeName);
        }

        // Ping failed — check OS mutex to distinguish "daemon busy" from "daemon not running".
        // The daemon holds this mutex for its entire lifetime, so its presence means
        // the daemon is running but temporarily unresponsive (e.g., during a heavy refresh).
        // This prevents starting a duplicate daemon (and a duplicate tray icon).
        if (IsDaemonMutexHeld(pipeName))
        {
            // Daemon is running but busy — wait briefly for it to become responsive.
            // If the mutex is still held after the wait window, do NOT try to start a
            // second daemon: it will immediately exit because the existing process still
            // owns the mutex. Surface an actionable recovery error instead.
            var waitUntil = DateTime.UtcNow + BusyDaemonWaitTimeout;
            while (DateTime.UtcNow < waitUntil)
            {
                await Task.Delay(BusyDaemonRetryInterval, cancellationToken);

                // Re-check mutex: if the daemon exited while we waited, stop waiting
                if (!IsDaemonMutexHeld(pipeName))
                    break;

                var remaining = waitUntil - DateTime.UtcNow;
                if (remaining <= TimeSpan.Zero)
                    break;

                if (await PingAsync(pipeName, Min(remaining, BusyDaemonConnectTimeout), cancellationToken))
                    return new ServiceClient(pipeName);
            }

            if (IsDaemonMutexHeld(pipeName))
            {
                throw new TimeoutException(
                    $"Daemon is running but not responding after {FormatDuration(BusyDaemonWaitTimeout)}. " +
                    "Stop it with 'excelcli service stop' or terminate the stuck excelcli process, then retry.");
            }

            // Daemon exited while we waited — start a replacement.
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
    internal static bool IsDaemonMutexHeld(string pipeName)
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
        var exePath = ResolveDaemonExecutablePath();

        var startInfo = new ProcessStartInfo
        {
            FileName = exePath,
            Arguments = $"service run --pipe-name \"{pipeName}\"",
            UseShellExecute = true,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden,
            WorkingDirectory = Path.GetDirectoryName(exePath) ?? Environment.CurrentDirectory
        };

        try
        {
            using var daemonProcess = Process.Start(startInfo)
                ?? throw new InvalidOperationException($"Failed to start daemon process '{exePath}'.");

            // Wait for daemon to be ready.
            var waitUntil = DateTime.UtcNow + StartupReadyTimeout;
            while (DateTime.UtcNow < waitUntil)
            {
                await Task.Delay(StartupReadyRetryInterval, cancellationToken);
                if (daemonProcess.HasExited)
                {
                    throw new InvalidOperationException(
                        $"Daemon process exited before becoming ready (exit code {daemonProcess.ExitCode}).");
                }

                var remaining = waitUntil - DateTime.UtcNow;
                if (remaining <= TimeSpan.Zero)
                    break;

                if (await PingAsync(pipeName, Min(remaining, StartupReadyConnectTimeout), cancellationToken))
                {
                    GC.KeepAlive(daemonProcess);
                    return;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to start daemon: {ex.Message}", ex);
        }

        throw new TimeoutException($"Daemon started but not responding within {FormatDuration(StartupReadyTimeout)}.");
    }

    private static string ResolveDaemonExecutablePath()
    {
        var baseDirectoryCandidate = Path.Combine(AppContext.BaseDirectory, "excelcli.exe");
        if (File.Exists(baseDirectoryCandidate))
        {
            return baseDirectoryCandidate;
        }

        var processPath = Environment.ProcessPath;
        if (!string.IsNullOrWhiteSpace(processPath) && File.Exists(processPath))
        {
            return processPath;
        }

        throw new InvalidOperationException("Cannot determine executable path to start daemon.");
    }

    private static string FormatDuration(TimeSpan duration)
    {
        return duration.TotalSeconds >= 1
            ? $"{duration.TotalSeconds:0.#} seconds"
            : $"{duration.TotalMilliseconds:0} ms";
    }

    private static TimeSpan Min(TimeSpan left, TimeSpan right) => left <= right ? left : right;

    private static async Task<bool> PingAsync(string pipeName, TimeSpan connectTimeout, CancellationToken cancellationToken)
    {
        var requestTimeout = connectTimeout + TimeSpan.FromSeconds(1);
        using var client = new ServiceClient(pipeName, connectTimeout: connectTimeout, requestTimeout: requestTimeout);
        return await client.PingAsync(cancellationToken);
    }
}
