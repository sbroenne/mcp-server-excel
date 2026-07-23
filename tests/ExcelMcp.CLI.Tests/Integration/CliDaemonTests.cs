using System.Diagnostics;
using System.IO.Pipes;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Integration tests for the CLI daemon process (excelcli service run).
/// Verifies the daemon starts, accepts pipe connections, and shuts down cleanly.
/// These tests do NOT require Excel — they validate daemon infrastructure.
/// Uses a test-specific pipe name to avoid conflicting with ServiceFixture.
/// </summary>
[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "ServiceDaemon")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Medium")]
public sealed class CliDaemonTests : IAsyncLifetime
{
    private readonly ITestOutputHelper _output;
    private readonly string _testPipeName = $"excelmcp-test-daemon-{Guid.NewGuid():N}";
    private Process? _daemonProcess;

    public CliDaemonTests(ITestOutputHelper output) => _output = output;

    public Task InitializeAsync()
    {
        // No need to stop existing daemons — we use a unique test pipe name
        return Task.CompletedTask;
    }

    public Task DisposeAsync()
    {
        KillDaemon();
        return Task.CompletedTask;
    }

    private Dictionary<string, string> TestEnv => new() { ["EXCELMCP_CLI_PIPE"] = _testPipeName };

    [Fact]
    public async Task ServiceStart_AutoStartsDaemonAndAcceptsConnections()
    {
        using var startProcess = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = CliProcessHelper.GetExePath(),
                Arguments = "-q service start",
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = Path.GetDirectoryName(CliProcessHelper.GetExePath())!
            }
        };
        foreach (var (key, value) in TestEnv)
        {
            startProcess.StartInfo.Environment[key] = value;
        }

        startProcess.Start();
        var exited = startProcess.WaitForExit(20000);
        Assert.True(exited, "service start should exit promptly after spawning the daemon");
        Assert.Equal(0, startProcess.ExitCode);

        var (statusResult, statusJson) = await CliProcessHelper.RunJsonAsync("service status", environmentVariables: TestEnv);
        _output.WriteLine($"Status response: {statusResult.Stdout}");

        Assert.Equal(0, statusResult.ExitCode);
        Assert.True(statusJson.RootElement.GetProperty("running").GetBoolean());

        var processId = statusJson.RootElement.GetProperty("processId").GetInt32();
        _daemonProcess = Process.GetProcessById(processId);
    }

    [Fact]
    public async Task ServiceStart_WhenDaemonMutexHeldButUnresponsive_ReturnsActionableError()
    {
        await using var heldMutex = await HeldMutex.AcquireAsync(DaemonAutoStart.GetDaemonMutexName(_testPipeName));

        var result = await CliProcessHelper.RunAsync("service start", timeoutMs: 20000, environmentVariables: TestEnv);
        _output.WriteLine($"Start response: {result.Stdout}");
        _output.WriteLine($"Start stderr: {result.Stderr}");

        using var json = JsonDocument.Parse(result.Stdout);
        Assert.Equal(1, result.ExitCode);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());

        var error = json.RootElement.GetProperty("error").GetString();
        Assert.NotNull(error);
        Assert.Contains("not responding", error, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("service stop", error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ServiceStart_WhenDaemonAcceptsConnectionButNeverReplies_ReturnsActionableError()
    {
        await using var heldMutex = await HeldMutex.AcquireAsync(DaemonAutoStart.GetDaemonMutexName(_testPipeName));

        await using var stalledDaemon = new StalledPipeServer(_testPipeName, _output);
        await stalledDaemon.StartAsync();

        var result = await CliProcessHelper.RunAsync("service start", timeoutMs: 20000, environmentVariables: TestEnv);
        _output.WriteLine($"Start response: {result.Stdout}");
        _output.WriteLine($"Start stderr: {result.Stderr}");

        using var json = JsonDocument.Parse(result.Stdout);
        Assert.Equal(1, result.ExitCode);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());

        var error = json.RootElement.GetProperty("error").GetString();
        Assert.NotNull(error);
        Assert.Contains("not responding", error, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("service stop", error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ServiceStart_WhenDaemonMutexExistsButIsNotOwned_StartsDaemon()
    {
        var mutexName = DaemonAutoStart.GetDaemonMutexName(_testPipeName);
        using var staleMutex = new Mutex(initiallyOwned: false, mutexName, out var createdNew);
        Assert.True(createdNew, "Test pipe should not have a pre-existing daemon mutex");

        var result = await CliProcessHelper.RunAsync("service start", timeoutMs: 15000, environmentVariables: TestEnv);
        _output.WriteLine($"Start response: {result.Stdout}");
        _output.WriteLine($"Start stderr: {result.Stderr}");

        using var startJson = JsonDocument.Parse(result.Stdout);
        Assert.Equal(0, result.ExitCode);
        Assert.True(startJson.RootElement.GetProperty("success").GetBoolean());

        var (statusResult, statusJson) = await CliProcessHelper.RunJsonAsync("service status", environmentVariables: TestEnv);
        Assert.Equal(0, statusResult.ExitCode);
        Assert.True(statusJson.RootElement.GetProperty("running").GetBoolean());

        var processId = statusJson.RootElement.GetProperty("processId").GetInt32();
        _daemonProcess = Process.GetProcessById(processId);
    }

    [Fact]
    public async Task ServiceStart_WhenStartupMutexIsAbandoned_StartsDaemon()
    {
        using var abandonedStartupMutex = CreateAbandonedMutex(DaemonAutoStart.GetDaemonStartupLockName(_testPipeName));

        var result = await CliProcessHelper.RunAsync("service start", timeoutMs: 15000, environmentVariables: TestEnv);
        _output.WriteLine($"Start response: {result.Stdout}");
        _output.WriteLine($"Start stderr: {result.Stderr}");

        using var startJson = JsonDocument.Parse(result.Stdout);
        Assert.Equal(0, result.ExitCode);
        Assert.True(startJson.RootElement.GetProperty("success").GetBoolean());

        var (statusResult, statusJson) = await CliProcessHelper.RunJsonAsync("service status", environmentVariables: TestEnv);
        Assert.Equal(0, statusResult.ExitCode);
        Assert.True(statusJson.RootElement.GetProperty("running").GetBoolean());

        var processId = statusJson.RootElement.GetProperty("processId").GetInt32();
        _daemonProcess = Process.GetProcessById(processId);
    }

    [Fact]
    public async Task ServiceStart_ConcurrentStartRequests_LeaveSingleResponsiveDaemon()
    {
        var startTasks = Enumerable.Range(0, 5)
            .Select(_ => CliProcessHelper.RunAsync("service start", timeoutMs: 20000, environmentVariables: TestEnv))
            .ToArray();

        var results = await Task.WhenAll(startTasks);
        foreach (var result in results)
        {
            _output.WriteLine($"Concurrent start response: {result.Stdout}");
            _output.WriteLine($"Concurrent start stderr: {result.Stderr}");
            Assert.Equal(0, result.ExitCode);
            using var json = JsonDocument.Parse(result.Stdout);
            Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        }

        var (statusResult, statusJson) = await CliProcessHelper.RunJsonAsync("service status", environmentVariables: TestEnv);
        Assert.Equal(0, statusResult.ExitCode);
        Assert.True(statusJson.RootElement.GetProperty("running").GetBoolean());

        var processId = statusJson.RootElement.GetProperty("processId").GetInt32();
        _daemonProcess = Process.GetProcessById(processId);
    }

    [Fact]
    public async Task ServiceStatus_WhenDaemonAcceptsConnectionButNeverReplies_FailsQuicklyWithTimeoutError()
    {
        await using var stalledDaemon = new StalledPipeServer(_testPipeName, _output);
        await stalledDaemon.StartAsync();

        var result = await CliProcessHelper.RunAsync("service status", timeoutMs: 10000, environmentVariables: TestEnv);
        _output.WriteLine($"Status response: {result.Stdout}");
        _output.WriteLine($"Status stderr: {result.Stderr}");

        using var json = JsonDocument.Parse(result.Stdout);
        Assert.Equal(1, result.ExitCode);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());
        Assert.False(json.RootElement.GetProperty("running").GetBoolean());

        var error = json.RootElement.GetProperty("error").GetString();
        Assert.NotNull(error);
        Assert.Contains("timed out", error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task SessionList_WhenDaemonAcceptsConnectionButNeverReplies_FailsQuicklyWithTimeoutError()
    {
        await using var stalledDaemon = new StalledPipeServer(_testPipeName, _output);
        await stalledDaemon.StartAsync();

        var result = await CliProcessHelper.RunAsync("session list", timeoutMs: 10000, environmentVariables: TestEnv);
        _output.WriteLine($"Session list response: {result.Stdout}");
        _output.WriteLine($"Session list stderr: {result.Stderr}");

        using var json = JsonDocument.Parse(result.Stdout);
        Assert.Equal(1, result.ExitCode);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());

        var error = json.RootElement.GetProperty("error").GetString();
        Assert.NotNull(error);
        Assert.Contains("timed out", error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ServiceRun_StartsAndAcceptsConnections()
    {
        // Start daemon as background process
        _daemonProcess = StartDaemon();
        _output.WriteLine($"Daemon started with PID {_daemonProcess.Id}, pipe: {_testPipeName}");

        // Wait for daemon pipe to be ready
        await WaitForDaemonReadyAsync();

        // Verify we can connect and get status
        var (result, json) = await CliProcessHelper.RunJsonAsync("service status", environmentVariables: TestEnv);
        _output.WriteLine($"Status response: {result.Stdout}");

        Assert.Equal(0, result.ExitCode);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        Assert.True(json.RootElement.GetProperty("running").GetBoolean());
        Assert.True(json.RootElement.GetProperty("processId").GetInt32() > 0);
    }

    [Fact]
    public async Task ServiceRun_ReportsZeroSessionsInitially()
    {
        _daemonProcess = StartDaemon();
        await WaitForDaemonReadyAsync();

        var (result, json) = await CliProcessHelper.RunJsonAsync("service status", environmentVariables: TestEnv);
        _output.WriteLine($"Status response: {result.Stdout}");

        Assert.Equal(0, result.ExitCode);
        Assert.Equal(0, json.RootElement.GetProperty("sessionCount").GetInt32());
    }

    [Fact]
    public async Task ServiceRun_AcceptsDiagPing()
    {
        _daemonProcess = StartDaemon();
        await WaitForDaemonReadyAsync();

        var (result, json) = await CliProcessHelper.RunJsonAsync("diag ping", environmentVariables: TestEnv);
        _output.WriteLine($"Ping response: {result.Stdout}");

        Assert.Equal(0, result.ExitCode);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("pong", json.RootElement.GetProperty("message").GetString());
    }

    [Fact]
    public async Task ServiceStop_ShutsDaemonDown()
    {
        _daemonProcess = StartDaemon();
        await WaitForDaemonReadyAsync();

        // Send stop command
        var stopResult = await CliProcessHelper.RunAsync("service stop", environmentVariables: TestEnv);
        _output.WriteLine($"Stop response: {stopResult.Stdout}");
        Assert.Equal(0, stopResult.ExitCode);

        // Wait for daemon process to exit
        var exited = _daemonProcess.WaitForExit(TimeSpan.FromSeconds(10));
        Assert.True(exited, "Daemon process should exit after stop command");
    }

    [Fact]
    public async Task ServiceStop_WhenDaemonIsUnresponsiveButTracked_ForceStopsTrackedProcess()
    {
        var mutexName = DaemonAutoStart.GetDaemonMutexName(_testPipeName);
        using var sleeper = StartMutexHoldingProcess(mutexName);
        await WaitForMutexAsync(mutexName);

        DaemonProcessTracker.RegisterProcess(
            _testPipeName,
            sleeper.Id,
            sleeper.StartTime.ToUniversalTime().ToFileTimeUtc());

        var result = await CliProcessHelper.RunAsync("service stop", timeoutMs: 15000, environmentVariables: TestEnv);
        _output.WriteLine($"Forced stop response: {result.Stdout}");
        _output.WriteLine($"Forced stop stderr: {result.Stderr}");

        using var json = JsonDocument.Parse(result.Stdout);
        Assert.Equal(0, result.ExitCode);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        Assert.True(json.RootElement.GetProperty("forced").GetBoolean());

        var exited = sleeper.WaitForExit(5000);
        Assert.True(exited, "Tracked daemon process should be force-stopped when shutdown RPC cannot get through");
        Assert.False(File.Exists(DaemonProcessTracker.GetTrackingFilePath(_testPipeName)));
    }

    [Fact]
    public async Task ServiceRun_SecondInstance_ExitsImmediatelyWithoutDuplicate()
    {
        // Start first daemon and wait until it is ready
        _daemonProcess = StartDaemon();
        await WaitForDaemonReadyAsync();
        _output.WriteLine($"First daemon running (PID {_daemonProcess.Id})");

        // Start a second daemon with the same pipe name — it should detect the mutex
        // held by the first daemon and exit immediately (exit code 0)
        var secondDaemon = StartDaemon();
        _output.WriteLine($"Second daemon started (PID {secondDaemon.Id})");

        var secondExited = secondDaemon.WaitForExit(TimeSpan.FromSeconds(5));
        _output.WriteLine(secondExited
            ? $"Second daemon exited with code {secondDaemon.ExitCode}"
            : "Second daemon did NOT exit within timeout — duplicate running!");

        try
        {
            Assert.True(secondExited, "Second daemon should exit immediately when a daemon is already running");
            Assert.Equal(0, secondDaemon.ExitCode);
        }
        finally
        {
            if (!secondDaemon.HasExited)
                secondDaemon.Kill(entireProcessTree: true);
            secondDaemon.Dispose();
        }

        // First daemon should still be alive and responsive
        var (statusResult, statusJson) = await CliProcessHelper.RunJsonAsync("service status", environmentVariables: TestEnv);
        Assert.Equal(0, statusResult.ExitCode);
        Assert.True(statusJson.RootElement.GetProperty("running").GetBoolean(),
            "First daemon should still be running after second-instance exit");
    }

    [Fact]
    public async Task ServiceRun_MutexReleasedAfterShutdown_NewDaemonCanStart()
    {
        // Start a daemon and shut it down
        var firstDaemon = StartDaemon();
        await WaitForDaemonReadyAsync();
        _output.WriteLine($"First daemon running (PID {firstDaemon.Id})");

        var stopResult = await CliProcessHelper.RunAsync("service stop", environmentVariables: TestEnv);
        Assert.Equal(0, stopResult.ExitCode);

        var firstExited = firstDaemon.WaitForExit(TimeSpan.FromSeconds(10));
        Assert.True(firstExited, "First daemon should exit after stop");
        firstDaemon.Dispose();
        _output.WriteLine("First daemon stopped");

        // A new daemon should now be able to start (mutex was released)
        _daemonProcess = StartDaemon();
        await WaitForDaemonReadyAsync();

        var (statusResult, statusJson) = await CliProcessHelper.RunJsonAsync("service status", environmentVariables: TestEnv);
        _output.WriteLine($"Second daemon status: {statusResult.Stdout}");

        Assert.Equal(0, statusResult.ExitCode);
        Assert.True(statusJson.RootElement.GetProperty("running").GetBoolean(),
            "A new daemon should start successfully after the previous one released the mutex");
    }

    private Process StartDaemon()
    {
        var exePath = CliProcessHelper.GetExePath();
        var startInfo = new ProcessStartInfo
        {
            FileName = exePath,
            Arguments = $"service run --pipe-name {_testPipeName}",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            WorkingDirectory = Path.GetDirectoryName(exePath)!
        };

        var process = new Process { StartInfo = startInfo };
        process.Start();
        return process;
    }

    private static Mutex CreateAbandonedMutex(string mutexName)
    {
        Mutex? mutex = null;
        Exception? threadException = null;
        using var acquired = new ManualResetEventSlim(false);
        var thread = new Thread(() =>
        {
            try
            {
                mutex = new Mutex(initiallyOwned: false, mutexName, out _);
                mutex.WaitOne();
                acquired.Set();
            }
            catch (Exception ex)
            {
                threadException = ex;
                acquired.Set();
            }
        })
        {
            IsBackground = true,
            Name = $"AbandonedMutex-{mutexName}"
        };

        thread.Start();
        Assert.True(acquired.Wait(TimeSpan.FromSeconds(5)), $"Test mutex '{mutexName}' was not acquired");
        Assert.True(thread.Join(TimeSpan.FromSeconds(5)), $"Test thread for mutex '{mutexName}' did not exit");
        Assert.Null(threadException);

        return mutex ?? throw new InvalidOperationException($"Test mutex '{mutexName}' was not created.");
    }

    private static Process StartMutexHoldingProcess(string mutexName)
    {
        var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = "powershell",
                Arguments = $"-NoLogo -NoProfile -NonInteractive -Command \"$created = $false; $m = New-Object System.Threading.Mutex($true, '{mutexName}', [ref]$created); if (-not $created) {{ exit 99 }}; try {{ Start-Sleep -Seconds 60 }} finally {{ $m.ReleaseMutex(); $m.Dispose() }}\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            }
        };

        process.Start();
        return process;
    }

    private static async Task WaitForMutexAsync(string mutexName, int maxRetries = 20, int delayMs = 250)
    {
        for (var i = 0; i < maxRetries; i++)
        {
            try
            {
                using var mutex = Mutex.OpenExisting(mutexName);
                return;
            }
            catch (WaitHandleCannotBeOpenedException)
            {
                await Task.Delay(delayMs);
            }
        }

        throw new TimeoutException($"Mutex '{mutexName}' was not acquired within {maxRetries * delayMs}ms");
    }

    private async Task WaitForDaemonReadyAsync(int maxRetries = 20, int delayMs = 500)
    {
        for (var i = 0; i < maxRetries; i++)
        {
            try
            {
                var result = await CliProcessHelper.RunAsync("service status", timeoutMs: 5000, environmentVariables: TestEnv);
                if (result.ExitCode == 0 && result.Stdout.Contains("\"running\":true"))
                {
                    _output.WriteLine($"Daemon ready after {(i + 1) * delayMs}ms");
                    return;
                }
            }
            catch (Exception)
            {
                // Daemon not ready yet
            }

            await Task.Delay(delayMs);
        }

        throw new TimeoutException($"CLI daemon did not become ready within {maxRetries * delayMs}ms");
    }

    private void KillDaemon()
    {
        if (_daemonProcess is null || _daemonProcess.HasExited) return;

        try
        {
            _daemonProcess.Kill(entireProcessTree: true);
            _daemonProcess.WaitForExit(TimeSpan.FromSeconds(5));
            _output.WriteLine($"Killed daemon PID {_daemonProcess.Id}");
        }
        catch (Exception ex)
        {
            _output.WriteLine($"Failed to kill daemon: {ex.Message}");
        }
        finally
        {
            _daemonProcess.Dispose();
        }
    }

    private sealed class StalledPipeServer : IAsyncDisposable
    {
        private readonly string _pipeName;
        private readonly ITestOutputHelper _output;
        private readonly CancellationTokenSource _cts = new();
        private readonly TaskCompletionSource _listeningTcs = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly Task _serverTask;
        private NamedPipeServerStream? _server;

        public StalledPipeServer(string pipeName, ITestOutputHelper output)
        {
            _pipeName = pipeName;
            _output = output;
            _serverTask = RunAsync();
        }

        public Task StartAsync() => _listeningTcs.Task;

        public async ValueTask DisposeAsync()
        {
            _cts.Cancel();

            if (_server != null)
            {
                try
                {
                    await _server.DisposeAsync();
                }
                catch
                {
                }
            }

            try
            {
                await _serverTask;
            }
            catch (OperationCanceledException)
            {
            }

            _cts.Dispose();
        }

        private async Task RunAsync()
        {
            try
            {
                _server = ServiceSecurity.CreateSecureServer(_pipeName);
                _listeningTcs.TrySetResult();
                _output.WriteLine($"Stalled pipe server listening on {_pipeName}");

                await _server.WaitForConnectionAsync(_cts.Token);
                _output.WriteLine("Stalled pipe server accepted a connection");

                await Task.Delay(Timeout.InfiniteTimeSpan, _cts.Token);
            }
            catch (OperationCanceledException)
            {
            }
            finally
            {
                if (_server != null)
                {
                    try
                    {
                        if (_server.IsConnected)
                        {
                            _server.Disconnect();
                        }
                    }
                    catch
                    {
                    }

                    await _server.DisposeAsync();
                    _server = null;
                }
            }
        }
    }

    private sealed class HeldMutex : IAsyncDisposable
    {
        private readonly string _mutexName;
        private readonly TaskCompletionSource _acquiredTcs = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly ManualResetEventSlim _releaseSignal = new(false);
        private readonly Thread _thread;
        private Mutex? _mutex;

        private HeldMutex(string mutexName)
        {
            _mutexName = mutexName;
            _thread = new Thread(ThreadMain)
            {
                IsBackground = true,
                Name = $"HeldMutex-{mutexName}"
            };
            _thread.Start();
        }

        public static async Task<HeldMutex> AcquireAsync(string mutexName)
        {
            var heldMutex = new HeldMutex(mutexName);
            await heldMutex._acquiredTcs.Task;
            return heldMutex;
        }

        public ValueTask DisposeAsync()
        {
            _releaseSignal.Set();
            _thread.Join(TimeSpan.FromSeconds(5));
            _releaseSignal.Dispose();
            return ValueTask.CompletedTask;
        }

        private void ThreadMain()
        {
            try
            {
                _mutex = new Mutex(initiallyOwned: true, _mutexName, out bool createdNew);
                if (!createdNew)
                {
                    throw new InvalidOperationException($"Failed to acquire unique test mutex '{_mutexName}'.");
                }

                _acquiredTcs.TrySetResult();
                _releaseSignal.Wait();
                _mutex.ReleaseMutex();
            }
            catch (Exception ex)
            {
                _acquiredTcs.TrySetException(ex);
            }
            finally
            {
                _mutex?.Dispose();
            }
        }
    }
}
