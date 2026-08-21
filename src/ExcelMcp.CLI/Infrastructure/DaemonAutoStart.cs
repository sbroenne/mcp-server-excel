using System.ComponentModel;
using System.Diagnostics;
using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>
/// Ensures the CLI daemon is running before sending commands.
/// Auto-starts the daemon if not already running.
/// </summary>
internal static class DaemonAutoStart
{
    internal static readonly TimeSpan InitialPingTimeout = DaemonConnectionPolicy.InitialProbeTimeout;
    internal static readonly TimeSpan BusyDaemonConnectTimeout = TimeSpan.FromSeconds(3);
    internal static readonly TimeSpan BusyDaemonRetryInterval = TimeSpan.FromMilliseconds(500);
    internal static readonly TimeSpan BusyDaemonWaitTimeout = TimeSpan.FromSeconds(10);
    internal static readonly TimeSpan StartupReadyConnectTimeout = TimeSpan.FromSeconds(1);
    internal static readonly TimeSpan StartupReadyRetryInterval = TimeSpan.FromMilliseconds(250);
    internal static readonly TimeSpan StartupReadyTimeout = DaemonConnectionPolicy.StartupReadyTimeout;
    internal static readonly TimeSpan StartupLockTimeout = DaemonStartupLock.Timeout;

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
        var startupDeadline = OperationDeadline.Start(StartupReadyTimeout);
        using var startupCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        startupCts.CancelAfter(StartupReadyTimeout);

        try
        {
            return await EnsureAndConnectCoreAsync(
                pipeName,
                startupDeadline,
                CreateRuntime(pipeName),
                startupCts.Token);
        }
        catch (OperationCanceledException) when (
            startupCts.IsCancellationRequested
            && !cancellationToken.IsCancellationRequested)
        {
            throw new TimeoutException(
                $"Daemon did not become ready within {FormatDuration(StartupReadyTimeout)}.");
        }
    }

    internal static async Task<ServiceClient> EnsureAndConnectCoreAsync(
        string pipeName,
        OperationDeadline startupDeadline,
        Runtime runtime,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(runtime);

        if (await runtime.PingAsync(
            startupDeadline.Cap(InitialPingTimeout),
            cancellationToken))
        {
            return new ServiceClient(pipeName);
        }

        if (runtime.IsDaemonMutexHeld())
        {
            var startupInProgress = RecheckStartupAfterDaemonObservation(
                startupAlreadyObserved: false,
                runtime.IsStartupInProgress);
            var responsivenessDeadline = startupInProgress
                ? startupDeadline
                : OperationDeadline.Start(startupDeadline.Cap(BusyDaemonWaitTimeout));
            while (true)
            {
                if (!startupInProgress && runtime.IsStartupInProgress())
                {
                    startupInProgress = true;
                    responsivenessDeadline = startupDeadline;
                }

                if (responsivenessDeadline.IsExpired)
                {
                    break;
                }

                await Task.Delay(
                    responsivenessDeadline.Cap(BusyDaemonRetryInterval),
                    cancellationToken);

                if (!runtime.IsDaemonMutexHeld())
                {
                    break;
                }

                var connectTimeout = responsivenessDeadline.Cap(BusyDaemonConnectTimeout);
                if (connectTimeout <= TimeSpan.Zero)
                {
                    break;
                }

                if (await runtime.PingAsync(connectTimeout, cancellationToken))
                {
                    return new ServiceClient(pipeName);
                }
            }

            if (!startupInProgress && runtime.IsStartupInProgress())
            {
                startupInProgress = true;
                if (await runtime.WaitForResponsiveDaemonAsync(
                    startupDeadline,
                    cancellationToken))
                {
                    return new ServiceClient(pipeName);
                }
            }
            var daemonStillRunning = runtime.IsDaemonMutexHeld();
            startupInProgress = RecheckStartupAfterDaemonObservation(
                startupInProgress,
                runtime.IsStartupInProgress);
            if (ShouldContinueStartupWait(
                    startupInProgress,
                    daemonStillRunning,
                    startupDeadline.IsExpired)
                && await runtime.WaitForResponsiveDaemonAsync(
                    startupDeadline,
                    cancellationToken))
            {
                return new ServiceClient(pipeName);
            }

            if (daemonStillRunning)
            {
                if (startupInProgress)
                {
                    throw new TimeoutException(
                        $"Daemon startup did not become ready within {FormatDuration(StartupReadyTimeout)}.");
                }

                throw new TimeoutException(
                    $"Daemon is running but not responding after {FormatDuration(BusyDaemonWaitTimeout)}. " +
                    "Stop it with 'excelcli service stop' or terminate the stuck excelcli process, then retry.");
            }
        }

        var startOutcome = await runtime.TryStartDaemonAsync(
            startupDeadline,
            cancellationToken);
        if (startOutcome == StartOutcome.Ready)
        {
            return new ServiceClient(pipeName);
        }

        if (await runtime.WaitForResponsiveDaemonAsync(
            startupDeadline,
            cancellationToken))
        {
            return new ServiceClient(pipeName);
        }

        throw new TimeoutException(startOutcome switch
        {
            StartOutcome.LockUnavailable =>
                $"CLI daemon lifecycle operation did not complete and no responsive daemon became ready within {FormatDuration(StartupReadyTimeout)}.",
            StartOutcome.ObserveReadiness =>
                $"Daemon started but not responding within {FormatDuration(StartupReadyTimeout)}.",
            _ => throw new InvalidOperationException($"Unexpected daemon start outcome '{startOutcome}'.")
        });
    }

    private static Runtime CreateRuntime(string pipeName) =>
        new(
            (timeout, cancellationToken) => PingAsync(pipeName, timeout, cancellationToken),
            () => IsDaemonMutexHeld(pipeName),
            () => IsDaemonStartupInProgress(pipeName),
            (deadline, cancellationToken) =>
                TryStartDaemonWithStartupLockAsync(pipeName, deadline, cancellationToken),
            (deadline, cancellationToken) =>
                WaitForResponsiveDaemonAsync(pipeName, deadline, cancellationToken));

    internal sealed record Runtime(
        Func<TimeSpan, CancellationToken, Task<bool>> PingAsync,
        Func<bool> IsDaemonMutexHeld,
        Func<bool> IsStartupInProgress,
        Func<OperationDeadline, CancellationToken, Task<StartOutcome>> TryStartDaemonAsync,
        Func<OperationDeadline, CancellationToken, Task<bool>> WaitForResponsiveDaemonAsync);

    internal enum StartOutcome
    {
        LockUnavailable,
        ObserveReadiness,
        Ready
    }

    internal static bool RecheckStartupAfterDaemonObservation(
        bool startupAlreadyObserved,
        Func<bool> startupMarkerProbe)
    {
        ArgumentNullException.ThrowIfNull(startupMarkerProbe);
        var startupObservedAfterDaemon = startupMarkerProbe();
        return startupAlreadyObserved || startupObservedAfterDaemon;
    }

    internal static bool ShouldContinueStartupWait(
        bool startupObserved,
        bool daemonObserved,
        bool startupDeadlineExpired) =>
        startupObserved
        && daemonObserved
        && !startupDeadlineExpired;

    /// <summary>
    /// Checks whether a daemon process currently holds the daemon mutex for the given pipe name.
    /// Returns true if a daemon is running (even if temporarily busy).
    /// </summary>
    internal static bool IsDaemonMutexHeld(string pipeName)
    {
        return IsMutexHeld(GetDaemonMutexName(pipeName))
            || DaemonStartupLock.GetLegacyDaemonMutexNames(pipeName)
                .Any(IsMutexHeld);
    }

    /// <summary>
    /// Checks whether a CLI process currently holds the active-startup marker.
    /// </summary>
    internal static bool IsDaemonStartupInProgress(string pipeName)
    {
        return IsMutexHeld(GetDaemonStartingMarkerName(pipeName));
    }

    private static bool IsMutexHeld(string mutexName)
    {
        Mutex? mutex = null;
        try
        {
            mutex = Mutex.OpenExisting(mutexName);
            if (mutex.WaitOne(TimeSpan.Zero))
            {
                mutex.ReleaseMutex();
                return false;
            }

            return true;
        }
        catch (AbandonedMutexException)
        {
            try { mutex?.ReleaseMutex(); } catch (ApplicationException) { }
            return false;
        }
        catch (WaitHandleCannotBeOpenedException)
        {
            return false;
        }
        catch (IOException)
        {
            return false;
        }
        catch (ArgumentException)
        {
            return false;
        }
        finally
        {
            mutex?.Dispose();
        }
    }

    /// <summary>
    /// Gets the OS mutex name for the CLI daemon identified by its pipe name.
    /// Used by both the daemon (to acquire) and the client (to detect a running daemon).
    /// </summary>
    internal static string GetDaemonMutexName(string pipeName) =>
        DaemonStartupLock.GetDaemonMutexName(pipeName);

    internal static string GetDaemonStartupLockName(string pipeName) =>
        DaemonStartupLock.GetStartupMutexName(pipeName);

    internal static string GetDaemonStartingMarkerName(string pipeName) =>
        DaemonStartupLock.GetStartingMarkerName(pipeName);

    internal static Task<T> WithStartupLockAsync<T>(
        string pipeName,
        Func<Task<T>> action,
        CancellationToken cancellationToken)
    {
        return DaemonStartupLock.WithLockAsync(
            pipeName,
            action,
            cancellationToken);
    }

    private static Task<StartOutcome> TryStartDaemonWithStartupLockAsync(
        string pipeName,
        OperationDeadline startupDeadline,
        CancellationToken cancellationToken)
    {
        return Task.Run(() =>
        {
            using var startupMutex = new Mutex(initiallyOwned: false, GetDaemonStartupLockName(pipeName), out _);
            Mutex? startingMarker = null;
            var startupLockAcquired = false;
            var startingMarkerAcquired = false;
            try
            {
                try
                {
                    var lockTimeout = startupDeadline.Cap(StartupLockTimeout);
                    startupLockAcquired = lockTimeout > TimeSpan.Zero
                        && startupMutex.WaitOne(lockTimeout);
                }
                catch (AbandonedMutexException)
                {
                    startupLockAcquired = true;
                }

                if (!startupLockAcquired)
                    return StartOutcome.LockUnavailable;

                // Another CLI process may have started the daemon while this process waited.
                var pingTimeout = startupDeadline.Cap(InitialPingTimeout);
                if (pingTimeout > TimeSpan.Zero
                    && PingAsync(pipeName, pingTimeout, cancellationToken).GetAwaiter().GetResult())
                {
                    return StartOutcome.Ready;
                }

                if (IsDaemonMutexHeld(pipeName))
                {
                    return StartOutcome.ObserveReadiness;
                }

                startingMarker = new Mutex(
                    initiallyOwned: false,
                    GetDaemonStartingMarkerName(pipeName),
                    out _);
                try
                {
                    startingMarkerAcquired = startingMarker.WaitOne(TimeSpan.Zero);
                }
                catch (AbandonedMutexException)
                {
                    startingMarkerAcquired = true;
                }

                if (!startingMarkerAcquired)
                {
                    throw new InvalidOperationException(
                        $"Could not acquire the daemon starting marker for CLI pipe '{pipeName}'.");
                }

                var cleanupResult = OwnedProcessCleanup
                    .CleanupAsync(pipeName, cancellationToken)
                    .GetAwaiter()
                    .GetResult();
                if (!cleanupResult.Success)
                {
                    throw new InvalidOperationException(
                        $"{cleanupResult.ErrorMessage ?? $"Tracked processes for CLI pipe '{pipeName}' could not be stopped."} " +
                        "Run 'excelcli service stop' and retry.");
                }

                // No daemon running — start it.
                StartDaemonAsync(
                    pipeName,
                    startupDeadline,
                    cancellationToken).GetAwaiter().GetResult();
                return StartOutcome.Ready;
            }
            finally
            {
                if (startingMarkerAcquired)
                {
                    startingMarker!.ReleaseMutex();
                }
                startingMarker?.Dispose();

                if (startupLockAcquired)
                {
                    startupMutex.ReleaseMutex();
                }
            }
        }, cancellationToken);
    }

    private static async Task<bool> WaitForResponsiveDaemonAsync(
        string pipeName,
        OperationDeadline startupDeadline,
        CancellationToken cancellationToken)
    {
        while (!startupDeadline.IsExpired)
        {
            await Task.Delay(
                startupDeadline.Cap(StartupReadyRetryInterval),
                cancellationToken);

            var connectTimeout = startupDeadline.Cap(StartupReadyConnectTimeout);
            if (connectTimeout <= TimeSpan.Zero)
            {
                break;
            }

            if (await PingAsync(pipeName, connectTimeout, cancellationToken))
            {
                return true;
            }
        }

        return false;
    }

    private static async Task StartDaemonAsync(
        string pipeName,
        OperationDeadline startupDeadline,
        CancellationToken cancellationToken)
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

        var daemonProcess = StartDaemonProcess(
            startupDeadline,
            cancellationToken,
            exePath,
            () => Process.Start(startInfo));

        using (daemonProcess)
        {
            while (!startupDeadline.IsExpired)
            {
                await Task.Delay(
                    startupDeadline.Cap(StartupReadyRetryInterval),
                    cancellationToken);
                if (daemonProcess.HasExited)
                {
                    if (daemonProcess.ExitCode == 0)
                    {
                        if (await WaitForResponsiveDaemonAsync(
                            pipeName,
                            startupDeadline,
                            cancellationToken))
                        {
                            GC.KeepAlive(daemonProcess);
                            return;
                        }

                        throw new InvalidOperationException(
                            "Daemon process exited cleanly before becoming ready, but no responsive daemon was found. " +
                            "This usually means a stale startup race or a daemon that shut down immediately. " +
                            "Run 'excelcli service stop' and retry.");
                    }

                    throw new InvalidOperationException(
                        $"Daemon process exited before becoming ready (exit code {daemonProcess.ExitCode}).");
                }

                var connectTimeout = startupDeadline.Cap(StartupReadyConnectTimeout);
                if (connectTimeout <= TimeSpan.Zero)
                {
                    break;
                }

                if (await PingAsync(pipeName, connectTimeout, cancellationToken))
                {
                    GC.KeepAlive(daemonProcess);
                    return;
                }
            }
        }

        throw new TimeoutException($"Daemon started but not responding within {FormatDuration(StartupReadyTimeout)}.");
    }

    internal static Process StartDaemonProcess(
        OperationDeadline startupDeadline,
        CancellationToken cancellationToken,
        string executablePath,
        Func<Process?> processStarter)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(executablePath);
        ArgumentNullException.ThrowIfNull(processStarter);
        cancellationToken.ThrowIfCancellationRequested();
        if (startupDeadline.IsExpired)
        {
            throw new TimeoutException(
                $"Daemon startup deadline expired before process '{executablePath}' could be launched.");
        }

        try
        {
            return processStarter()
                ?? throw new InvalidOperationException($"Failed to start daemon process '{executablePath}'.");
        }
        catch (Win32Exception ex)
        {
            throw new InvalidOperationException(
                $"Failed to start daemon process '{executablePath}': {ex.Message}",
                ex);
        }
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

    private static async Task<bool> PingAsync(string pipeName, TimeSpan connectTimeout, CancellationToken cancellationToken)
    {
        if (connectTimeout <= TimeSpan.Zero)
        {
            return false;
        }

        using var client = new ServiceClient(
            pipeName,
            connectTimeout: connectTimeout,
            requestTimeout: connectTimeout);
        var response = await client.SendAsync(
            new ServiceRequest { Command = "service.ping" },
            connectTimeout,
            cancellationToken);
        return response.Success;
    }
}
