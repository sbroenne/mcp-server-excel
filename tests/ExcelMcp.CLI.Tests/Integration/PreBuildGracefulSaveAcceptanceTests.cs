using System.Diagnostics;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "ServiceDaemon")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Slow")]
public sealed class PreBuildGracefulSaveAcceptanceTests : IClassFixture<TempDirectoryFixture>
{
    private readonly TempDirectoryFixture _fixture;

    public PreBuildGracefulSaveAcceptanceTests(TempDirectoryFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact(Timeout = 360000)]
    public async Task StaleLockedBuildCleanup_GracefullySavesDirtySession_AndPreservesOtherPipe()
    {
        var repositoryRoot = GetRepositoryRoot();
        var releaseDirectory = Path.Combine(
            repositoryRoot,
            "src",
            "ExcelMcp.CLI",
            "bin",
            "Release",
            "net10.0-windows");
        var releaseCli = Path.Combine(releaseDirectory, "excelcli.exe");
        Assert.True(File.Exists(releaseCli), $"Release CLI is required: {releaseCli}");

        var controlDirectory = Path.Combine(
            _fixture.TempDir,
            $"control-cli-{Guid.NewGuid():N}");
        CopyDirectory(releaseDirectory, controlDirectory);
        var controlCli = Path.Combine(controlDirectory, "excelcli.exe");
        var selectedPipe = $"excelmcp-stale-save-{Guid.NewGuid():N}";
        var controlPipe = $"excelmcp-stale-save-control-{Guid.NewGuid():N}";
        var workbookPath = Path.Combine(
            _fixture.TempDir,
            $"stale-save-{Guid.NewGuid():N}.xlsx");
        const string marker = "graceful-stale-cleanup-persisted";
        using var selectedDaemon = StartDaemonProcess(releaseCli, selectedPipe);
        using var controlDaemon = StartDaemonProcess(controlCli, controlPipe);
        var safetySource = Path.Combine(
            repositoryRoot,
            "src",
            "ExcelMcp.Service",
            "ServiceClient.cs");
        var originalSourceWriteTime = File.GetLastWriteTimeUtc(safetySource);

        try
        {
            await WaitForDaemonReadyAsync(releaseCli, selectedPipe);
            await WaitForDaemonReadyAsync(controlCli, controlPipe);

            var create = await RunCliAsync(
                releaseCli,
                selectedPipe,
                ["session", "create", workbookPath, "--quiet"],
                TimeSpan.FromSeconds(45));
            Assert.Equal(0, create.ExitCode);
            using var createJson = JsonDocument.Parse(create.Stdout);
            var sessionId = createJson.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrWhiteSpace(sessionId));

            var write = await RunCliAsync(
                releaseCli,
                selectedPipe,
                [
                    "range", "set-values",
                    "--session", sessionId!,
                    "--sheet-name", "Sheet1",
                    "--range-address", "A1",
                    "--values", JsonSerializer.Serialize(new[] { new[] { marker } }),
                    "--quiet"
                ],
                TimeSpan.FromSeconds(30));
            Assert.Equal(0, write.ExitCode);

            File.SetLastWriteTimeUtc(safetySource, DateTime.UtcNow.AddSeconds(2));
            Assert.True(
                File.GetLastWriteTimeUtc(safetySource) > File.GetLastWriteTimeUtc(releaseCli),
                "The current protocol source must be newer than the locked CLI.");

            var build = await RunProcessAsync(
                "dotnet",
                [
                    "build",
                    Path.Combine(
                        repositoryRoot,
                        "src",
                        "ExcelMcp.CLI",
                        "ExcelMcp.CLI.csproj"),
                    "--configuration", "Release",
                    "-p:NuGetAudit=false",
                    "-maxcpucount:1",
                    "-nodeReuse:false",
                    "--verbosity", "quiet"
                ],
                repositoryRoot,
                new Dictionary<string, string>
                {
                    ["EXCELMCP_CLI_PIPE"] = selectedPipe
                },
                TimeSpan.FromMinutes(4));

            Assert.True(
                build.ExitCode == 0,
                $"Release rebuild failed.{Environment.NewLine}" +
                $"stdout: {build.Stdout}{Environment.NewLine}" +
                $"stderr: {build.Stderr}");
            Assert.True(
                selectedDaemon.WaitForExit(15000),
                $"Selected daemon {selectedDaemon.Id} did not exit.");
            Assert.False(controlDaemon.HasExited);
            Assert.Equal(0, (await GetServiceStatusAsync(controlCli, controlPipe)).ExitCode);

            var reopen = await RunCliAsync(
                releaseCli,
                selectedPipe,
                ["session", "open", workbookPath, "--quiet"],
                TimeSpan.FromSeconds(45));
            Assert.Equal(0, reopen.ExitCode);
            using var reopenJson = JsonDocument.Parse(reopen.Stdout);
            var reopenedSessionId = reopenJson.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrWhiteSpace(reopenedSessionId));

            var read = await RunCliAsync(
                releaseCli,
                selectedPipe,
                [
                    "range", "get-values",
                    "--session", reopenedSessionId!,
                    "--sheet-name", "Sheet1",
                    "--range-address", "A1",
                    "--quiet"
                ],
                TimeSpan.FromSeconds(30));
            Assert.Equal(0, read.ExitCode);
            using var readJson = JsonDocument.Parse(read.Stdout);
            Assert.Equal(marker, readJson.RootElement.GetProperty("values")[0][0].GetString());

            _ = await RunCliAsync(
                releaseCli,
                selectedPipe,
                ["session", "close", "--session", reopenedSessionId!, "--quiet"],
                TimeSpan.FromSeconds(45));
        }
        finally
        {
            File.SetLastWriteTimeUtc(safetySource, originalSourceWriteTime);
            await StopDaemonBestEffortAsync(releaseCli, selectedPipe, selectedDaemon);
            await StopDaemonBestEffortAsync(controlCli, controlPipe, controlDaemon);
        }
    }

    private static Process StartDaemonProcess(string cliPath, string pipeName)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = cliPath,
            WorkingDirectory = Path.GetDirectoryName(cliPath)!,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        startInfo.ArgumentList.Add("service");
        startInfo.ArgumentList.Add("run");
        startInfo.ArgumentList.Add("--pipe-name");
        startInfo.ArgumentList.Add(pipeName);
        startInfo.ArgumentList.Add("--quiet");
        return Process.Start(startInfo)!;
    }

    private static async Task WaitForDaemonReadyAsync(string cliPath, string pipeName)
    {
        var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(20);
        while (DateTime.UtcNow < deadline)
        {
            try
            {
                var status = await GetServiceStatusAsync(cliPath, pipeName);
                if (status.ExitCode == 0
                    && status.Stdout.Contains("\"running\":true", StringComparison.Ordinal))
                {
                    return;
                }
            }
            catch (TimeoutException) when (DateTime.UtcNow < deadline)
            {
                // A bounded status probe can time out while the daemon is still starting.
            }

            await Task.Delay(250);
        }

        throw new TimeoutException($"Daemon for pipe '{pipeName}' did not become ready.");
    }

    private static Task<ProcessResult> GetServiceStatusAsync(string cliPath, string pipeName) =>
        RunCliAsync(
            cliPath,
            pipeName,
            ["service", "status", "--quiet"],
            DaemonConnectionPolicy.ControlTimeout + TimeSpan.FromSeconds(2));

    private static Task<ProcessResult> RunCliAsync(
        string cliPath,
        string pipeName,
        IReadOnlyList<string> arguments,
        TimeSpan timeout) =>
        RunProcessAsync(
            cliPath,
            arguments,
            Path.GetDirectoryName(cliPath)!,
            new Dictionary<string, string>
            {
                ["EXCELMCP_CLI_PIPE"] = pipeName
            },
            timeout);

    private static async Task StopDaemonBestEffortAsync(
        string cliPath,
        string pipeName,
        Process daemon)
    {
        _ = await RunCliAsync(
            cliPath,
            pipeName,
            ["service", "stop", "--quiet"],
            TimeSpan.FromSeconds(20));

        if (!daemon.HasExited)
        {
            daemon.Kill(entireProcessTree: true);
        }

        daemon.WaitForExit(5000);
    }

    private static async Task<ProcessResult> RunProcessAsync(
        string fileName,
        IReadOnlyList<string> arguments,
        string workingDirectory,
        IReadOnlyDictionary<string, string>? environmentVariables,
        TimeSpan timeout)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = fileName,
            WorkingDirectory = workingDirectory,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        foreach (var argument in arguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        if (environmentVariables != null)
        {
            foreach (var (key, value) in environmentVariables)
            {
                startInfo.Environment[key] = value;
            }
        }

        using var process = Process.Start(startInfo)!;
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        using var timeoutCts = new CancellationTokenSource(timeout);
        try
        {
            await process.WaitForExitAsync(timeoutCts.Token);
        }
        catch (OperationCanceledException)
        {
            process.Kill(entireProcessTree: true);
            await process.WaitForExitAsync();
            throw new TimeoutException(
                $"Process '{fileName}' timed out after {timeout}.");
        }

        return new ProcessResult(
            process.ExitCode,
            await stdoutTask,
            await stderrTask);
    }

    private static void CopyDirectory(string sourceDirectory, string destinationDirectory)
    {
        Directory.CreateDirectory(destinationDirectory);
        foreach (var file in Directory.EnumerateFiles(sourceDirectory))
        {
            File.Copy(file, Path.Combine(destinationDirectory, Path.GetFileName(file)));
        }

        foreach (var directory in Directory.EnumerateDirectories(sourceDirectory))
        {
            CopyDirectory(
                directory,
                Path.Combine(destinationDirectory, Path.GetFileName(directory)));
        }
    }

    private static string GetRepositoryRoot() =>
        Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));

    private sealed record ProcessResult(int ExitCode, string Stdout, string Stderr);
}
