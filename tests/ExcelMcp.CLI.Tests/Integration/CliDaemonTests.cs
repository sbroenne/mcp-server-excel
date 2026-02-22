using System.Diagnostics;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
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
}
