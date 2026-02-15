using System.Diagnostics;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Integration tests for the CLI daemon process (excelcli service run).
/// Verifies the daemon starts, accepts pipe connections, and shuts down cleanly.
/// These tests do NOT require Excel — they validate daemon infrastructure.
/// </summary>
[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "ServiceDaemon")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Medium")]
public sealed class CliDaemonTests : IAsyncLifetime
{
    private readonly ITestOutputHelper _output;
    private Process? _daemonProcess;

    public CliDaemonTests(ITestOutputHelper output) => _output = output;

    public async Task InitializeAsync()
    {
        // Kill any leftover daemon that might hold the CLI pipe
        await StopExistingDaemonAsync();
    }

    public Task DisposeAsync()
    {
        KillDaemon();
        return Task.CompletedTask;
    }

    [Fact]
    public async Task ServiceRun_StartsAndAcceptsConnections()
    {
        // Start daemon as background process
        _daemonProcess = StartDaemon();
        _output.WriteLine($"Daemon started with PID {_daemonProcess.Id}");

        // Wait for daemon pipe to be ready
        await WaitForDaemonReadyAsync();

        // Verify we can connect and get status
        var (result, json) = await CliProcessHelper.RunJsonAsync("service status");
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

        var (result, json) = await CliProcessHelper.RunJsonAsync("service status");
        _output.WriteLine($"Status response: {result.Stdout}");

        Assert.Equal(0, result.ExitCode);
        Assert.Equal(0, json.RootElement.GetProperty("sessionCount").GetInt32());
    }

    [Fact]
    public async Task ServiceRun_AcceptsDiagPing()
    {
        _daemonProcess = StartDaemon();
        await WaitForDaemonReadyAsync();

        var (result, json) = await CliProcessHelper.RunJsonAsync("diag ping");
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
        var stopResult = await CliProcessHelper.RunAsync("service stop");
        _output.WriteLine($"Stop response: {stopResult.Stdout}");
        Assert.Equal(0, stopResult.ExitCode);

        // Wait for daemon process to exit
        var exited = _daemonProcess.WaitForExit(TimeSpan.FromSeconds(10));
        Assert.True(exited, "Daemon process should exit after stop command");
    }

    private static Process StartDaemon()
    {
        var exePath = CliProcessHelper.GetExePath();
        var startInfo = new ProcessStartInfo
        {
            FileName = exePath,
            Arguments = "service run",
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
                var result = await CliProcessHelper.RunAsync("service status", timeoutMs: 5000);
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

    private async Task StopExistingDaemonAsync()
    {
        try
        {
            var result = await CliProcessHelper.RunAsync("service stop", timeoutMs: 5000);
            if (result.ExitCode == 0)
            {
                _output.WriteLine("Stopped existing daemon");
                await Task.Delay(1000); // Give it time to fully exit
            }
        }
        catch (Exception)
        {
            // No existing daemon — fine
        }
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
