using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
using System.Xml.Linq;
using Sbroenne.ExcelMcp.CLI.Commands;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Regression coverage for pipe-scoped daemon and orphaned Excel process cleanup.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "CLI")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
public sealed class DaemonForcedStopRegressionTests
{
    public static TheoryData<string> AdversarialPipeNames
    {
        get
        {
            var data = new TheoryData<string>();
            foreach (var pipeName in GetAdversarialPipeNames())
            {
                data.Add(pipeName);
            }

            return data;
        }
    }

    [Fact]
    public void MutexNames_UseDisjointHashedNamespaces()
    {
        const string daemonPrefix = "ExcelMcpCli_Daemon_";
        const string startupPrefix = "ExcelMcpCli_Startup_";
        const string trackerPrefix = "ExcelMcpCli_Tracker_";
        var pipeNames = GetAdversarialPipeNames();
        var names = pipeNames
            .SelectMany(pipeName => new[]
            {
                DaemonAutoStart.GetDaemonMutexName(pipeName),
                DaemonAutoStart.GetDaemonStartupLockName(pipeName),
                DaemonProcessTracker.GetTrackingMutexName(pipeName)
            })
            .ToList();

        Assert.Equal(
            pipeNames.Distinct(StringComparer.OrdinalIgnoreCase).Count() * 3,
            names.Distinct(StringComparer.Ordinal).Count());
        foreach (var pipeName in pipeNames)
        {
            var daemonName = DaemonAutoStart.GetDaemonMutexName(pipeName);
            var startupName = DaemonAutoStart.GetDaemonStartupLockName(pipeName);
            var trackerName = DaemonProcessTracker.GetTrackingMutexName(pipeName);
            Assert.StartsWith(daemonPrefix, daemonName, StringComparison.Ordinal);
            Assert.StartsWith(startupPrefix, startupName, StringComparison.Ordinal);
            Assert.StartsWith(trackerPrefix, trackerName, StringComparison.Ordinal);
            Assert.Equal(64, daemonName[daemonPrefix.Length..].Length);
            Assert.Equal(64, startupName[startupPrefix.Length..].Length);
            Assert.Equal(64, trackerName[trackerPrefix.Length..].Length);
            Assert.All(daemonName[daemonPrefix.Length..], character => Assert.True(char.IsAsciiHexDigit(character)));
            Assert.All(startupName[startupPrefix.Length..], character => Assert.True(char.IsAsciiHexDigit(character)));
            Assert.All(trackerName[trackerPrefix.Length..], character => Assert.True(char.IsAsciiHexDigit(character)));
            Assert.Equal(daemonName, DaemonAutoStart.GetDaemonMutexName(pipeName));
            Assert.Equal(startupName, DaemonAutoStart.GetDaemonStartupLockName(pipeName));
            Assert.Equal(trackerName, DaemonProcessTracker.GetTrackingMutexName(pipeName));
        }

        Assert.Equal(
            DaemonAutoStart.GetDaemonMutexName("foo"),
            DaemonAutoStart.GetDaemonMutexName("FOO"));
        Assert.Equal(
            DaemonAutoStart.GetDaemonStartupLockName("foo"),
            DaemonAutoStart.GetDaemonStartupLockName("FOO"));
        Assert.Equal(
            DaemonProcessTracker.GetTrackingMutexName("foo"),
            DaemonProcessTracker.GetTrackingMutexName("FOO"));
        Assert.Equal(
            DaemonProcessTracker.GetTrackingFilePath("foo"),
            DaemonProcessTracker.GetTrackingFilePath("FOO"));
    }

    [Fact(Timeout = 60000)]
    public async Task CaseVariantPipeNamesShareDaemonAndCleanupIdentity()
    {
        var cliPath = CliProcessHelper.GetExePath();
        var pipeName = $"excelmcp-case-foo-{Guid.NewGuid():N}";
        var caseVariantPipeName = pipeName.ToUpperInvariant();
        using var daemon = StartDaemonProcess(cliPath, pipeName);
        Process? duplicateDaemon = null;

        try
        {
            await WaitForDaemonReadyAsync(cliPath, pipeName);
            duplicateDaemon = StartDaemonProcess(cliPath, caseVariantPipeName);
            Assert.True(
                duplicateDaemon.WaitForExit(10000),
                "A case-variant pipe name must not start a second daemon.");

            var status = await GetServiceStatusAsync(cliPath, caseVariantPipeName);
            Assert.Equal(0, status.ExitCode);
            using var statusJson = JsonDocument.Parse(status.Stdout);
            Assert.Equal(
                daemon.Id,
                statusJson.RootElement.GetProperty("processId").GetInt32());

            var stopResult = await RunProcessAsync(
                cliPath,
                ["service", "stop", "--quiet"],
                Path.GetDirectoryName(cliPath)!,
                new Dictionary<string, string>
                {
                    ["EXCELMCP_CLI_PIPE"] = caseVariantPipeName
                },
                TimeSpan.FromSeconds(20));

            Assert.Equal(0, stopResult.ExitCode);
            Assert.True(daemon.WaitForExit(10000));
            Assert.Equal(
                DaemonProcessTracker.GetTrackingFilePath(pipeName),
                DaemonProcessTracker.GetTrackingFilePath(caseVariantPipeName));
            Assert.False(File.Exists(DaemonProcessTracker.GetTrackingFilePath(pipeName)));
        }
        finally
        {
            if (duplicateDaemon != null)
            {
                await StopDaemonBestEffortAsync(
                    cliPath,
                    caseVariantPipeName,
                    duplicateDaemon);
                duplicateDaemon.Dispose();
            }
            await StopDaemonBestEffortAsync(cliPath, pipeName, daemon);
            DaemonProcessTracker.Clear(pipeName);
            DaemonProcessTracker.Clear(caseVariantPipeName);
        }
    }

    [Fact(Timeout = 60000)]
    public async Task LegacyDaemonMutex_PreventsNewDaemonFromStarting()
    {
        var cliPath = CliProcessHelper.GetExePath();
        var pipeName = $"excelmcp-legacy-mutex-{Guid.NewGuid():N}";
        using var legacyMutex = new Mutex(
            initiallyOwned: true,
            $"ExcelMcpCli_{pipeName}",
            out var createdNew);
        Assert.True(createdNew);
        using var daemon = StartDaemonProcess(cliPath, pipeName);

        try
        {
            Assert.True(
                daemon.WaitForExit(5000),
                "A daemon must exit when a legacy daemon already owns the pipe lifetime mutex.");
            Assert.Equal(0, daemon.ExitCode);
        }
        finally
        {
            if (!daemon.HasExited)
            {
                daemon.Kill(entireProcessTree: true);
                await daemon.WaitForExitAsync();
            }
            DaemonProcessTracker.Clear(pipeName);
        }
    }

    [Fact(Timeout = 60000)]
    public async Task StartupMutex_DoesNotCollideWithSuffixNamedDaemonMutex()
    {
        var cliPath = CliProcessHelper.GetExePath();
        var pipeName = $"excelmcp-mutex-collision-{Guid.NewGuid():N}";
        var suffixPipeName = $"{pipeName}_startup";
        using var daemon = StartDaemonProcess(cliPath, pipeName);
        using var suffixDaemon = StartDaemonProcess(cliPath, suffixPipeName);

        try
        {
            await WaitForDaemonReadyAsync(cliPath, pipeName);
            await WaitForDaemonReadyAsync(cliPath, suffixPipeName);

            var stopResult = await RunProcessAsync(
                cliPath,
                ["service", "stop", "--quiet"],
                Path.GetDirectoryName(cliPath)!,
                new Dictionary<string, string>
                {
                    ["EXCELMCP_CLI_PIPE"] = pipeName
                },
                TimeSpan.FromSeconds(20));

            Assert.True(
                stopResult.ExitCode == 0,
                $"service stop failed with exit code {stopResult.ExitCode}.{Environment.NewLine}" +
                $"stdout: {stopResult.Stdout}{Environment.NewLine}stderr: {stopResult.Stderr}");
            Assert.True(daemon.WaitForExit(10000));
            Assert.False(suffixDaemon.HasExited);
            Assert.True((await GetServiceStatusAsync(cliPath, suffixPipeName)).ExitCode == 0);
        }
        finally
        {
            await StopDaemonBestEffortAsync(cliPath, suffixPipeName, suffixDaemon);
            await StopDaemonBestEffortAsync(cliPath, pipeName, daemon);
        }
    }

    [Theory]
    [MemberData(nameof(AdversarialPipeNames))]
    public async Task ServiceStop_AdversarialPipeNamesRemainIsolated(string pipeName)
    {
        var controlPipeName = $"{pipeName}-control";
        using var daemon = StartSleepingProcess();
        using var excel = StartSleepingProcess();
        using var controlDaemon = StartSleepingProcess();
        using var controlExcel = StartSleepingProcess();

        try
        {
            RegisterTrackedProcesses(pipeName, daemon, excel);
            RegisterTrackedProcesses(controlPipeName, controlDaemon, controlExcel);

            var result = await RunServiceStopAsync(pipeName);

            Assert.Equal(0, result.ExitCode);
            Assert.True(daemon.WaitForExit(5000));
            Assert.True(excel.WaitForExit(5000));
            Assert.False(controlDaemon.HasExited);
            Assert.False(controlExcel.HasExited);
            Assert.True(File.Exists(DaemonProcessTracker.GetTrackingFilePath(controlPipeName)));
        }
        finally
        {
            StopIfRunning(daemon);
            StopIfRunning(excel);
            StopIfRunning(controlDaemon);
            StopIfRunning(controlExcel);
            DaemonProcessTracker.Clear(pipeName);
            DaemonProcessTracker.Clear(controlPipeName);
        }
    }

    [Fact]
    public async Task StartupMutex_SamePipeSerializesCallers()
    {
        var pipeName = $"excelmcp-same-pipe-lock-{Guid.NewGuid():N}";
        var firstEntered = new TaskCompletionSource(
            TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseFirst = new TaskCompletionSource(
            TaskCreationOptions.RunContinuationsAsynchronously);
        var secondEntered = new TaskCompletionSource(
            TaskCreationOptions.RunContinuationsAsynchronously);

        var first = DaemonStartupLock.WithLockAsync(
            pipeName,
            async () =>
            {
                firstEntered.SetResult();
                await releaseFirst.Task;
                return true;
            },
            CancellationToken.None);
        await firstEntered.Task.WaitAsync(TimeSpan.FromSeconds(5));

        var second = DaemonStartupLock.WithLockAsync(
            pipeName,
            () =>
            {
                secondEntered.SetResult();
                return Task.FromResult(true);
            },
            CancellationToken.None);

        await Task.Delay(250);
        Assert.False(secondEntered.Task.IsCompleted);

        releaseFirst.SetResult();
        Assert.True(await first);
        Assert.True(await second);
        Assert.True(secondEntered.Task.IsCompleted);
    }

    [Fact]
    public void PreBuildCleanup_RunsOnlyForCliProject()
    {
        var buildProperties = XDocument.Load(
            Path.Combine(GetRepositoryRoot(), "Directory.Build.props"));
        var cleanupTarget = buildProperties
            .Descendants("Target")
            .Single(element => string.Equals(
                element.Attribute("Name")?.Value,
                "StopExcelMcpProcesses",
                StringComparison.Ordinal));

        var condition = cleanupTarget.Attribute("Condition")?.Value;

        Assert.Contains(
            "'$(MSBuildProjectName)' == 'ExcelMcp.CLI'",
            condition,
            StringComparison.Ordinal);
        Assert.Equal("BeforeBuild", cleanupTarget.Attribute("BeforeTargets")?.Value);
    }

    [Fact]
    public void CleanupScript_StagesCurrentClientWhenExistingBinaryIsStale()
    {
        var cleanupScript = File.ReadAllText(
            Path.Combine(GetRepositoryRoot(), "scripts", "Stop-ExcelMcpProcesses.ps1"));

        Assert.Contains(
            @"src\ExcelMcp.CLI\Infrastructure\DaemonAutoStart.cs",
            cleanupScript,
            StringComparison.Ordinal);
        Assert.Contains(
            @"src\ExcelMcp.CLI\Program.cs",
            cleanupScript,
            StringComparison.Ordinal);
        Assert.Contains(
            @"src\ExcelMcp.ComInterop\Session\SessionManager.cs",
            cleanupScript,
            StringComparison.Ordinal);
        Assert.Contains(
            @"src\ExcelMcp.ComInterop\Session\ExcelBatch.cs",
            cleanupScript,
            StringComparison.Ordinal);
        Assert.Contains(
            @"src\ExcelMcp.ComInterop\Session\ExcelProcessIdentity.cs",
            cleanupScript,
            StringComparison.Ordinal);
        Assert.Contains(
            @"src\ExcelMcp.ComInterop\Session\OwnedProcessGuard.cs",
            cleanupScript,
            StringComparison.Ordinal);
        Assert.Contains(
            "-p:ExcelMcpCleanupRoot=$stagingRoot",
            cleanupScript,
            StringComparison.Ordinal);
        Assert.Contains(
            "-p:ExcelMcpSkipCleanup=true",
            cleanupScript,
            StringComparison.Ordinal);
        Assert.Contains(
            @"src\ExcelMcp.Cleanup\ExcelMcp.Cleanup.csproj",
            cleanupScript,
            StringComparison.Ordinal);
        Assert.Contains(
            @"src\ExcelMcp.Service\ServiceClient.cs",
            cleanupScript,
            StringComparison.Ordinal);
        Assert.Contains(
            @"src\ExcelMcp.Service\Rpc\IExcelDaemonRpc.cs",
            cleanupScript,
            StringComparison.Ordinal);
        Assert.Contains(
            "if ($availableClis.Count -eq 0)",
            cleanupScript,
            StringComparison.Ordinal);
    }

    [Fact(Timeout = 300000)]
    public async Task PreBuildCleanup_StaleBinaryStopsOwnedDaemonAndAllowsReleaseRebuild()
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
        Assert.True(File.Exists(releaseCli), $"Release CLI is required for this regression: {releaseCli}");

        var controlDirectory = Path.Combine(
            Path.GetTempPath(),
            $"excelmcp-stale-cleanup-control-{Guid.NewGuid():N}");
        CopyDirectory(releaseDirectory, controlDirectory);
        var controlCli = Path.Combine(controlDirectory, "excelcli.exe");
        var ownedPipe = $"excelmcp-stale-build-owned-{Guid.NewGuid():N}";
        var controlPipe = $"excelmcp-stale-build-control-{Guid.NewGuid():N}";
        using var ownedDaemon = StartDaemonProcess(releaseCli, ownedPipe);
        using var controlDaemon = StartDaemonProcess(controlCli, controlPipe);
        var safetySource = Path.Combine(
            repositoryRoot,
            "src",
            "ExcelMcp.CLI",
            "Infrastructure",
            "DaemonProcessTracker.cs");
        var originalSourceWriteTime = File.GetLastWriteTimeUtc(safetySource);

        try
        {
            await WaitForDaemonReadyAsync(releaseCli, ownedPipe);
            await WaitForDaemonReadyAsync(controlCli, controlPipe);

            File.SetLastWriteTimeUtc(safetySource, DateTime.UtcNow);
            Assert.True(
                File.GetLastWriteTimeUtc(safetySource) > File.GetLastWriteTimeUtc(releaseCli),
                "The ownership source must be newer than the existing CLI for this regression.");

            var buildResult = await RunProcessAsync(
                "dotnet",
                [
                    "build",
                    Path.Combine(
                        repositoryRoot,
                        "src",
                        "ExcelMcp.CLI",
                        "ExcelMcp.CLI.csproj"),
                    "--configuration",
                    "Release",
                    "-p:NuGetAudit=false",
                    "-maxcpucount:1",
                    "-nodeReuse:false",
                    "--verbosity",
                    "quiet"
                ],
                repositoryRoot,
                new Dictionary<string, string>
                {
                    ["EXCELMCP_CLI_PIPE"] = ownedPipe
                },
                TimeSpan.FromMinutes(3));

            Assert.True(
                buildResult.ExitCode == 0,
                $"Release rebuild failed.{Environment.NewLine}" +
                $"stdout: {buildResult.Stdout}{Environment.NewLine}" +
                $"stderr: {buildResult.Stderr}");
            Assert.DoesNotContain("warning ", buildResult.Stdout, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("warning ", buildResult.Stderr, StringComparison.OrdinalIgnoreCase);
            Assert.True(
                ownedDaemon.WaitForExit(10000),
                $"Owned daemon {ownedDaemon.Id} was not stopped by the stale-binary pre-build target.");
            Assert.False(controlDaemon.HasExited);
            Assert.True((await GetServiceStatusAsync(controlCli, controlPipe)).ExitCode == 0);
        }
        finally
        {
            File.SetLastWriteTimeUtc(safetySource, originalSourceWriteTime);
            await StopDaemonBestEffortAsync(releaseCli, ownedPipe, ownedDaemon);
            await StopDaemonBestEffortAsync(controlCli, controlPipe, controlDaemon);
            Directory.Delete(controlDirectory, recursive: true);
        }
    }

    [Fact]
    public async Task Cleanup_StopsExcelTrackedAndUntrackedBetweenSnapshots()
    {
        var pipeName = $"excelmcp-between-snapshots-{Guid.NewGuid():N}";
        using var daemon = StartSleepingProcess();
        using var excel = StartSleepingProcess();

        try
        {
            var daemonIdentity = DaemonProcessTracker.RegisterProcess(
                pipeName,
                daemon.Id,
                daemon.StartTime.ToUniversalTime().ToFileTimeUtc());
            var preShutdownSnapshot = OwnedProcessCleanup.CaptureTrackedProcesses(pipeName);

            DaemonProcessTracker.UpdateExcelProcesses(pipeName, daemonIdentity, [excel.Id]);
            DaemonProcessTracker.UpdateExcelProcesses(pipeName, daemonIdentity, []);
            var trackedExcel = Assert.Single(
                DaemonProcessTracker.GetTrackedExcelProcessIdentities(pipeName));
            Assert.Equal(excel.Id, trackedExcel.ProcessId);

            var cleanupResult = await OwnedProcessCleanup.CleanupAsync(
                pipeName,
                preShutdownSnapshot,
                CancellationToken.None);

            Assert.True(cleanupResult.Success);
            Assert.True(daemon.WaitForExit(5000));
            Assert.True(excel.WaitForExit(5000));
            Assert.False(File.Exists(DaemonProcessTracker.GetTrackingFilePath(pipeName)));
        }
        finally
        {
            StopIfRunning(daemon);
            StopIfRunning(excel);
            DaemonProcessTracker.Clear(pipeName);
        }
    }

    [Fact]
    public async Task Cleanup_DelayedStaleExcelIdentityDoesNotCaptureReusedPid()
    {
        var pipeName = $"excelmcp-delayed-pid-reuse-{Guid.NewGuid():N}";
        using var daemon = StartSleepingProcess();
        using var replacement = StartSleepingProcess();

        try
        {
            var daemonIdentity = DaemonProcessTracker.RegisterProcess(
                pipeName,
                daemon.Id,
                daemon.StartTime.ToUniversalTime().ToFileTimeUtc());
            var preShutdownSnapshot = OwnedProcessCleanup.CaptureTrackedProcesses(pipeName);
            var staleExcelIdentity = new DaemonProcessTracker.ProcessIdentity(
                replacement.Id,
                replacement.StartTime.ToUniversalTime().ToFileTimeUtc() + 1);

            DaemonProcessTracker.RecordExcelProcesses(
                pipeName,
                daemonIdentity,
                [staleExcelIdentity]);

            var trackedExcel = Assert.Single(
                DaemonProcessTracker.GetTrackedExcelProcessIdentities(pipeName));
            Assert.Equal(staleExcelIdentity, trackedExcel);

            var cleanupResult = await OwnedProcessCleanup.CleanupAsync(
                pipeName,
                preShutdownSnapshot,
                CancellationToken.None);

            Assert.True(cleanupResult.Success);
            Assert.True(daemon.WaitForExit(5000));
            Assert.False(replacement.HasExited);
            Assert.False(File.Exists(DaemonProcessTracker.GetTrackingFilePath(pipeName)));
        }
        finally
        {
            StopIfRunning(daemon);
            StopIfRunning(replacement);
            DaemonProcessTracker.Clear(pipeName);
        }
    }

    [Fact]
    public async Task Cleanup_DoesNotStopUntrackedDaemonChildProcess()
    {
        var pipeName = $"excelmcp-child-isolation-{Guid.NewGuid():N}";
        var childPidFile = Path.Combine(Path.GetTempPath(), $"excelmcp-child-{Guid.NewGuid():N}.pid");
        using var daemon = StartSleepingProcessWithChild(childPidFile);
        using var child = await WaitForChildProcessAsync(childPidFile);

        try
        {
            DaemonProcessTracker.RegisterProcess(
                pipeName,
                daemon.Id,
                daemon.StartTime.ToUniversalTime().ToFileTimeUtc());

            var cleanupResult = await OwnedProcessCleanup.CleanupAsync(pipeName, CancellationToken.None);

            Assert.True(cleanupResult.Success);
            Assert.True(daemon.WaitForExit(5000));
            Assert.False(child.HasExited);
        }
        finally
        {
            StopIfRunning(daemon);
            StopIfRunning(child);
            DaemonProcessTracker.Clear(pipeName);
            File.Delete(childPidFile);
        }
    }

    [Fact]
    public async Task Cleanup_PreservesTrackedExcelWhenGracefulDaemonUntracksDuringShutdown()
    {
        var pipeName = $"excelmcp-graceful-tracking-loss-{Guid.NewGuid():N}";
        using var daemon = StartShortLivedProcess();
        using var excel = StartSleepingProcess();

        try
        {
            var daemonIdentity = RegisterTrackedProcesses(pipeName, daemon, excel);
            var preShutdownSnapshot = OwnedProcessCleanup.CaptureTrackedProcesses(pipeName);

            DaemonProcessTracker.UpdateExcelProcesses(pipeName, daemonIdentity, []);
            var trackedExcel = Assert.Single(
                DaemonProcessTracker.GetTrackedExcelProcessIdentities(pipeName));
            Assert.Equal(excel.Id, trackedExcel.ProcessId);
            await daemon.WaitForExitAsync();
            Assert.Equal(0, daemon.ExitCode);
            Assert.False(excel.HasExited);

            var cleanupResult = await OwnedProcessCleanup.CleanupAsync(
                pipeName,
                preShutdownSnapshot,
                CancellationToken.None);

            Assert.True(cleanupResult.Success);
            Assert.True(daemon.HasExited);
            Assert.True(excel.WaitForExit(5000));
            Assert.False(File.Exists(DaemonProcessTracker.GetTrackingFilePath(pipeName)));
        }
        finally
        {
            StopIfRunning(daemon);
            StopIfRunning(excel);
            DaemonProcessTracker.Clear(pipeName);
        }
    }

    [Fact]
    public async Task Cleanup_DoesNotStopOrClearReplacementDaemonGeneration()
    {
        var pipeName = $"excelmcp-replacement-generation-{Guid.NewGuid():N}";
        using var oldDaemon = StartSleepingProcess();
        using var oldExcel = StartSleepingProcess();
        using var replacementDaemon = StartSleepingProcess();
        using var replacementExcel = StartSleepingProcess();

        try
        {
            RegisterTrackedProcesses(pipeName, oldDaemon, oldExcel);
            var oldSnapshot = OwnedProcessCleanup.CaptureTrackedProcesses(pipeName);
            StopIfRunning(oldDaemon);

            RegisterTrackedProcesses(pipeName, replacementDaemon, replacementExcel);

            var cleanupResult = await OwnedProcessCleanup.CleanupAsync(
                pipeName,
                oldSnapshot,
                CancellationToken.None);

            Assert.True(cleanupResult.Success);
            Assert.True(oldExcel.WaitForExit(5000));
            Assert.False(replacementDaemon.HasExited);
            Assert.False(replacementExcel.HasExited);
            Assert.True(File.Exists(DaemonProcessTracker.GetTrackingFilePath(pipeName)));
        }
        finally
        {
            StopIfRunning(oldDaemon);
            StopIfRunning(oldExcel);
            StopIfRunning(replacementDaemon);
            StopIfRunning(replacementExcel);
            DaemonProcessTracker.Clear(pipeName);
        }
    }

    [Fact]
    public void UpdateExcelProcesses_DelayedOldGenerationDoesNotOverwriteReplacement()
    {
        var pipeName = $"excelmcp-delayed-tracking-{Guid.NewGuid():N}";
        using var oldDaemon = StartSleepingProcess();
        using var replacementDaemon = StartSleepingProcess();
        using var oldExcel = StartSleepingProcess();

        try
        {
            var oldDaemonIdentity = DaemonProcessTracker.RegisterProcess(
                pipeName,
                oldDaemon.Id,
                oldDaemon.StartTime.ToUniversalTime().ToFileTimeUtc());
            DaemonProcessTracker.RegisterProcess(
                pipeName,
                replacementDaemon.Id,
                replacementDaemon.StartTime.ToUniversalTime().ToFileTimeUtc());

            DaemonProcessTracker.UpdateExcelProcesses(
                pipeName,
                oldDaemonIdentity,
                [oldExcel.Id]);

            Assert.True(DaemonProcessTracker.TryGetProcessSnapshot(pipeName, out var snapshot));
            Assert.Equal(replacementDaemon.Id, snapshot.DaemonProcess.ProcessId);
            Assert.Empty(snapshot.ExcelProcesses);
        }
        finally
        {
            StopIfRunning(oldDaemon);
            StopIfRunning(replacementDaemon);
            StopIfRunning(oldExcel);
            DaemonProcessTracker.Clear(pipeName);
        }
    }

    [Fact]
    public async Task ServiceStop_CleansOnlyProcessesTrackedForSelectedPipe()
    {
        var pipeA = $"excelmcp-owned-cleanup-a-{Guid.NewGuid():N}";
        var pipeB = $"excelmcp-owned-cleanup-b-{Guid.NewGuid():N}";
        using var daemonA = StartSleepingProcess();
        using var excelA = StartSleepingProcess();
        using var daemonB = StartSleepingProcess();
        using var excelB = StartSleepingProcess();

        try
        {
            RegisterTrackedProcesses(pipeA, daemonA, excelA);
            RegisterTrackedProcesses(pipeB, daemonB, excelB);

            var result = await RunServiceStopAsync(pipeA);

            Assert.True(
                result.ExitCode == 0,
                $"service stop failed with exit code {result.ExitCode}.{Environment.NewLine}" +
                $"stdout: {result.Stdout}{Environment.NewLine}stderr: {result.Stderr}");
            Assert.True(daemonA.WaitForExit(5000));
            Assert.True(excelA.WaitForExit(5000));
            Assert.False(daemonB.HasExited);
            Assert.False(excelB.HasExited);
            Assert.False(File.Exists(DaemonProcessTracker.GetTrackingFilePath(pipeA)));
            Assert.True(File.Exists(DaemonProcessTracker.GetTrackingFilePath(pipeB)));
        }
        finally
        {
            StopIfRunning(daemonA);
            StopIfRunning(excelA);
            StopIfRunning(daemonB);
            StopIfRunning(excelB);
            DaemonProcessTracker.Clear(pipeA);
            DaemonProcessTracker.Clear(pipeB);
        }
    }

    [Fact]
    public async Task ServiceStop_CleansTrackedExcelAfterDaemonExited()
    {
        var pipeName = $"excelmcp-orphan-cleanup-{Guid.NewGuid():N}";
        using var daemon = StartSleepingProcess();
        using var excel = StartSleepingProcess();

        try
        {
            RegisterTrackedProcesses(pipeName, daemon, excel);
            StopIfRunning(daemon);

            var result = await RunServiceStopAsync(pipeName);

            Assert.Equal(0, result.ExitCode);
            Assert.True(excel.WaitForExit(5000));
            Assert.False(File.Exists(DaemonProcessTracker.GetTrackingFilePath(pipeName)));
        }
        finally
        {
            StopIfRunning(daemon);
            StopIfRunning(excel);
            DaemonProcessTracker.Clear(pipeName);
        }
    }

    [Fact]
    public async Task ServiceStop_IgnoresPidReusedTrackingRecords()
    {
        var pipeName = $"excelmcp-stale-cleanup-{Guid.NewGuid():N}";
        using var daemon = StartSleepingProcess();
        using var excel = StartSleepingProcess();

        try
        {
            WriteStaleTrackingRecord(pipeName, daemon, excel);

            var result = await RunServiceStopAsync(pipeName);

            Assert.True(
                result.ExitCode == 0,
                $"service stop failed with exit code {result.ExitCode}.{Environment.NewLine}" +
                $"stdout: {result.Stdout}{Environment.NewLine}stderr: {result.Stderr}");
            Assert.False(daemon.HasExited);
            Assert.False(excel.HasExited);
            Assert.False(File.Exists(DaemonProcessTracker.GetTrackingFilePath(pipeName)));
        }
        finally
        {
            StopIfRunning(daemon);
            StopIfRunning(excel);
            DaemonProcessTracker.Clear(pipeName);
        }
    }

    [Fact]
    public async Task ServiceStop_OwnedCleanupIsIdempotent()
    {
        var pipeName = $"excelmcp-idempotent-cleanup-{Guid.NewGuid():N}";
        using var daemon = StartSleepingProcess();
        using var excel = StartSleepingProcess();

        try
        {
            RegisterTrackedProcesses(pipeName, daemon, excel);

            var firstResult = await RunServiceStopAsync(pipeName);
            var secondResult = await RunServiceStopAsync(pipeName);

            Assert.Equal(0, firstResult.ExitCode);
            Assert.Equal(0, secondResult.ExitCode);
            Assert.True(daemon.WaitForExit(5000));
            Assert.True(excel.WaitForExit(5000));
            Assert.False(File.Exists(DaemonProcessTracker.GetTrackingFilePath(pipeName)));
        }
        finally
        {
            StopIfRunning(daemon);
            StopIfRunning(excel);
            DaemonProcessTracker.Clear(pipeName);
        }
    }

    [Fact]
    public async Task ForceStopTrackedDaemon_AlsoStopsTrackedExcelProcess()
    {
        var pipeName = $"excelmcp-force-stop-test-{Guid.NewGuid():N}";
        using var daemon = StartSleepingProcess();
        using var excel = StartSleepingProcess();

        try
        {
            var daemonIdentity = DaemonProcessTracker.RegisterProcess(
                pipeName,
                daemon.Id,
                daemon.StartTime.ToUniversalTime().ToFileTimeUtc());
            DaemonProcessTracker.UpdateExcelProcesses(pipeName, daemonIdentity, [excel.Id]);
            var preShutdownSnapshot = OwnedProcessCleanup.CaptureTrackedProcesses(pipeName);

            var method = typeof(ServiceStopCommand).GetMethod(
                "TryForceStopTrackedDaemonAsync",
                BindingFlags.Static | BindingFlags.NonPublic);
            Assert.NotNull(method);

            var stopTask = (Task<bool>)method!.Invoke(
                null,
                [pipeName, preShutdownSnapshot, CancellationToken.None])!;
            Assert.True(await stopTask);

            await daemon.WaitForExitAsync();
            await excel.WaitForExitAsync();
            Assert.True(daemon.HasExited);
            Assert.True(excel.HasExited);
        }
        finally
        {
            StopIfRunning(daemon);
            StopIfRunning(excel);
            DaemonProcessTracker.Clear(pipeName);
        }
    }

    [Fact]
    public void TerminateProcess_AlreadyExited_IsTreatedAsSuccess()
    {
        using var process = Process.Start(new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = "-NoProfile -Command \"exit 0\"",
            UseShellExecute = false,
            CreateNoWindow = true
        })!;
        process.WaitForExit();

        var terminated = OwnedProcessCleanup.TryTerminateProcess(process, entireProcessTree: false);
        Assert.True(terminated);
    }

    [Fact]
    public void TryGetTrackedProcess_TransientInvalidRecord_DoesNotDeleteTrackingFile()
    {
        var pipeName = $"excelmcp-tracker-read-test-{Guid.NewGuid():N}";
        var trackingFile = DaemonProcessTracker.GetTrackingFilePath(pipeName);

        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(trackingFile)!);
            File.WriteAllText(trackingFile, """{"processId":""");

            Assert.False(DaemonProcessTracker.TryGetTrackedProcess(pipeName, out _));
            Assert.True(File.Exists(trackingFile));
        }
        finally
        {
            DaemonProcessTracker.Clear(pipeName);
        }
    }

    [Fact]
    public void UpdateExcelProcesses_TransientInvalidRecord_DoesNotBreakLifecycle()
    {
        var pipeName = $"excelmcp-tracker-update-test-{Guid.NewGuid():N}";
        var trackingFile = DaemonProcessTracker.GetTrackingFilePath(pipeName);

        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(trackingFile)!);
            File.WriteAllText(trackingFile, """{"processId":""");
            using var currentProcess = Process.GetCurrentProcess();
            var daemonIdentity = new DaemonProcessTracker.ProcessIdentity(
                currentProcess.Id,
                currentProcess.StartTime.ToUniversalTime().ToFileTimeUtc());

            var exception = Record.Exception(() =>
                DaemonProcessTracker.UpdateExcelProcesses(
                    pipeName,
                    daemonIdentity,
                    [Environment.ProcessId]));

            Assert.Null(exception);
            Assert.True(File.Exists(trackingFile));
        }
        finally
        {
            DaemonProcessTracker.Clear(pipeName);
        }
    }

    [Fact]
    public void RecordExcelProcesses_InvalidRecord_FailsStrictPersistence()
    {
        var pipeName = $"excelmcp-tracker-strict-test-{Guid.NewGuid():N}";
        var trackingFile = DaemonProcessTracker.GetTrackingFilePath(pipeName);

        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(trackingFile)!);
            File.WriteAllText(trackingFile, """{"processId":""");
            using var currentProcess = Process.GetCurrentProcess();
            var daemonIdentity = new DaemonProcessTracker.ProcessIdentity(
                currentProcess.Id,
                currentProcess.StartTime.ToUniversalTime().ToFileTimeUtc());

            var exception = Assert.Throws<InvalidOperationException>(() =>
                DaemonProcessTracker.RecordExcelProcesses(
                    pipeName,
                    daemonIdentity,
                    [new DaemonProcessTracker.ProcessIdentity(
                        Environment.ProcessId,
                        currentProcess.StartTime.ToUniversalTime().ToFileTimeUtc())]));
            Assert.Contains("malformed", exception.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            DaemonProcessTracker.Clear(pipeName);
        }
    }

    [Fact]
    public async Task Cleanup_MissingTrackingRecord_IsSafeNoOp()
    {
        var pipeName = $"excelmcp-tracker-missing-{Guid.NewGuid():N}";

        var result = await OwnedProcessCleanup.CleanupAsync(
            pipeName,
            CancellationToken.None);

        Assert.True(result.Success);
        Assert.False(result.DaemonMatched);
    }

    [Fact]
    public async Task Cleanup_MalformedTrackingRecord_FailsAndPreservesEvidence()
    {
        var pipeName = $"excelmcp-tracker-malformed-{Guid.NewGuid():N}";
        var trackingFile = DaemonProcessTracker.GetTrackingFilePath(pipeName);

        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(trackingFile)!);
            await File.WriteAllTextAsync(trackingFile, """{"processId":""");

            var result = await OwnedProcessCleanup.CleanupAsync(
                pipeName,
                CancellationToken.None);

            Assert.False(result.Success);
            Assert.False(result.DaemonMatched);
            Assert.True(File.Exists(trackingFile));
            Assert.Equal("""{"processId":""", await File.ReadAllTextAsync(trackingFile));
        }
        finally
        {
            DaemonProcessTracker.Clear(pipeName);
        }
    }

    [Fact]
    public async Task Cleanup_UnreadableTrackingRecord_FailsAndPreservesEvidence()
    {
        var pipeName = $"excelmcp-tracker-unreadable-{Guid.NewGuid():N}";
        var trackingFile = DaemonProcessTracker.GetTrackingFilePath(pipeName);

        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(trackingFile)!);
            await File.WriteAllTextAsync(trackingFile, """{"processId":1}""");
            await using (var exclusiveLease = new FileStream(
                trackingFile,
                FileMode.Open,
                FileAccess.ReadWrite,
                FileShare.None))
            {
                var result = await OwnedProcessCleanup.CleanupAsync(
                    pipeName,
                    CancellationToken.None);

                Assert.False(result.Success);
                Assert.False(result.DaemonMatched);
                Assert.True(File.Exists(trackingFile));
            }
        }
        finally
        {
            DaemonProcessTracker.Clear(pipeName);
        }
    }

    [Fact]
    public async Task ServiceStop_MalformedTrackingRecord_ReportsFailureAndPreservesEvidence()
    {
        var pipeName = $"excelmcp-tracker-service-stop-{Guid.NewGuid():N}";
        var trackingFile = DaemonProcessTracker.GetTrackingFilePath(pipeName);

        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(trackingFile)!);
            await File.WriteAllTextAsync(trackingFile, """{"processId":""");

            var result = await RunServiceStopAsync(pipeName);

            Assert.Equal(1, result.ExitCode);
            using var response = JsonDocument.Parse(result.Stdout);
            Assert.False(response.RootElement.GetProperty("success").GetBoolean());
            Assert.Contains(
                "malformed",
                response.RootElement.GetProperty("error").GetString(),
                StringComparison.OrdinalIgnoreCase);
            Assert.True(File.Exists(trackingFile));
        }
        finally
        {
            DaemonProcessTracker.Clear(pipeName);
        }
    }

    [Fact(Timeout = 180000)]
    public async Task PreBuildCleanup_MalformedTrackingRecordPropagatesFailure()
    {
        var pipeName = $"excelmcp-tracker-prebuild-{Guid.NewGuid():N}";
        var trackingFile = DaemonProcessTracker.GetTrackingFilePath(pipeName);

        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(trackingFile)!);
            await File.WriteAllTextAsync(trackingFile, """{"processId":""");

            var result = await RunProcessAsync(
                "powershell.exe",
                [
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    Path.Combine(
                        GetRepositoryRoot(),
                        "scripts",
                        "Stop-ExcelMcpProcesses.ps1"),
                    "-PipeName",
                    pipeName,
                    "-Verbose"
                ],
                GetRepositoryRoot(),
                environmentVariables: null,
                TimeSpan.FromMinutes(2));

            Assert.Equal(1, result.ExitCode);
            Assert.Contains(
                "cleanup failed",
                result.Stdout + result.Stderr,
                StringComparison.OrdinalIgnoreCase);
            Assert.True(File.Exists(trackingFile));
        }
        finally
        {
            DaemonProcessTracker.Clear(pipeName);
        }
    }

    [Fact(Timeout = 180000)]
    public async Task PreBuildCleanup_UnresponsiveFallbackRemainsPipeScoped()
    {
        var selectedPipe = $"excelmcp-prebuild-unresponsive-{Guid.NewGuid():N}";
        var unrelatedPipe = $"excelmcp-prebuild-control-{Guid.NewGuid():N}";
        using var selectedDaemon = StartSleepingProcess();
        using var selectedExcel = StartSleepingProcess();
        using var unrelatedDaemon = StartSleepingProcess();
        using var unrelatedExcel = StartSleepingProcess();
        var repositoryRoot = GetRepositoryRoot();
        var protocolSource = Path.Combine(
            repositoryRoot,
            "src",
            "ExcelMcp.Service",
            "ServiceClient.cs");
        var originalSourceWriteTime = File.GetLastWriteTimeUtc(protocolSource);

        try
        {
            RegisterTrackedProcesses(selectedPipe, selectedDaemon, selectedExcel);
            RegisterTrackedProcesses(unrelatedPipe, unrelatedDaemon, unrelatedExcel);
            File.SetLastWriteTimeUtc(protocolSource, DateTime.UtcNow.AddSeconds(2));

            var result = await RunProcessAsync(
                "powershell.exe",
                [
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    Path.Combine(
                        repositoryRoot,
                        "scripts",
                        "Stop-ExcelMcpProcesses.ps1"),
                    "-PipeName",
                    selectedPipe,
                    "-Verbose"
                ],
                repositoryRoot,
                environmentVariables: null,
                TimeSpan.FromMinutes(2));

            Assert.Equal(0, result.ExitCode);
            Assert.True(selectedDaemon.WaitForExit(5000));
            Assert.True(selectedExcel.WaitForExit(5000));
            Assert.False(unrelatedDaemon.HasExited);
            Assert.False(unrelatedExcel.HasExited);
            Assert.True(File.Exists(DaemonProcessTracker.GetTrackingFilePath(unrelatedPipe)));
        }
        finally
        {
            File.SetLastWriteTimeUtc(protocolSource, originalSourceWriteTime);
            StopIfRunning(selectedDaemon);
            StopIfRunning(selectedExcel);
            StopIfRunning(unrelatedDaemon);
            StopIfRunning(unrelatedExcel);
            DaemonProcessTracker.Clear(selectedPipe);
            DaemonProcessTracker.Clear(unrelatedPipe);
        }
    }

    [Fact]
    public async Task Cleanup_LegacyCaseSensitiveRecordUsesCaseInsensitivePipeIdentity()
    {
        var legacyPipeName = $"excelmcp-legacy-tracker-{Guid.NewGuid():N}".ToLowerInvariant();
        var callerPipeName = legacyPipeName.ToUpperInvariant();
        var legacyTrackingFile = GetLegacyTrackingFilePath(legacyPipeName);
        using var daemon = StartSleepingProcess();
        using var excel = StartSleepingProcess();

        try
        {
            WriteTrackingRecord(legacyTrackingFile, daemon, excel);

            var result = await OwnedProcessCleanup.CleanupAsync(
                callerPipeName,
                CancellationToken.None);

            Assert.True(result.Success);
            Assert.True(result.DaemonMatched);
            Assert.True(daemon.WaitForExit(5000));
            Assert.True(excel.WaitForExit(5000));
            Assert.False(File.Exists(legacyTrackingFile));
        }
        finally
        {
            StopIfRunning(daemon);
            StopIfRunning(excel);
            DaemonProcessTracker.Clear(legacyPipeName);
            DaemonProcessTracker.Clear(callerPipeName);
            File.Delete(legacyTrackingFile);
        }
    }

    [Fact]
    public async Task Cleanup_LegacyCaseSensitiveRecordKeepsUnrelatedPipeIsolated()
    {
        var selectedPipe = $"excelmcp-legacy-selected-{Guid.NewGuid():N}".ToLowerInvariant();
        var unrelatedPipe = $"excelmcp-legacy-unrelated-{Guid.NewGuid():N}".ToLowerInvariant();
        var selectedTrackingFile = GetLegacyTrackingFilePath(selectedPipe);
        var unrelatedTrackingFile = GetLegacyTrackingFilePath(unrelatedPipe);
        using var selectedDaemon = StartSleepingProcess();
        using var selectedExcel = StartSleepingProcess();
        using var unrelatedDaemon = StartSleepingProcess();
        using var unrelatedExcel = StartSleepingProcess();

        try
        {
            WriteTrackingRecord(selectedTrackingFile, selectedDaemon, selectedExcel);
            WriteTrackingRecord(unrelatedTrackingFile, unrelatedDaemon, unrelatedExcel);

            var result = await OwnedProcessCleanup.CleanupAsync(
                selectedPipe.ToUpperInvariant(),
                CancellationToken.None);

            Assert.True(result.Success);
            Assert.True(selectedDaemon.WaitForExit(5000));
            Assert.True(selectedExcel.WaitForExit(5000));
            Assert.False(unrelatedDaemon.HasExited);
            Assert.False(unrelatedExcel.HasExited);
            Assert.True(File.Exists(unrelatedTrackingFile));
        }
        finally
        {
            StopIfRunning(selectedDaemon);
            StopIfRunning(selectedExcel);
            StopIfRunning(unrelatedDaemon);
            StopIfRunning(unrelatedExcel);
            DaemonProcessTracker.Clear(selectedPipe);
            DaemonProcessTracker.Clear(selectedPipe.ToUpperInvariant());
            DaemonProcessTracker.Clear(unrelatedPipe);
            File.Delete(selectedTrackingFile);
            File.Delete(unrelatedTrackingFile);
        }
    }

    [Fact]
    public async Task TrackingMutex_LegacyCaseVariantSerializesCanonicalCaller()
    {
        var legacyPipeName = $"excelmcp-legacy-lock-{Guid.NewGuid():N}".ToLowerInvariant();
        var callerPipeName = legacyPipeName.ToUpperInvariant();
        var legacyMutexName =
            $"ExcelMcpCliTracker_{DaemonPipeIdentity.GetCaseSensitiveHash(legacyPipeName)}";
        using var acquired = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var holder = Task.Run(() =>
        {
            using var legacyMutex = new Mutex(
                initiallyOwned: false,
                legacyMutexName,
                out var createdNew);
            Assert.True(createdNew);
            legacyMutex.WaitOne();
            try
            {
                acquired.Set();
                release.Wait();
            }
            finally
            {
                legacyMutex.ReleaseMutex();
            }
        });

        Assert.True(acquired.Wait(TimeSpan.FromSeconds(5)));
        var registration = Task.Run(() =>
            DaemonProcessTracker.RegisterProcess(
                callerPipeName,
                Environment.ProcessId,
                Process.GetCurrentProcess().StartTime.ToUniversalTime().ToFileTimeUtc()));
        try
        {
            await Task.Delay(250);
            Assert.False(registration.IsCompleted);
        }
        finally
        {
            release.Set();
            await holder;
            await registration;
            DaemonProcessTracker.Clear(callerPipeName);
        }
    }

    private static Process StartSleepingProcess()
    {
        return Process.Start(new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = "-NoProfile -Command \"Start-Sleep -Seconds 60\"",
            UseShellExecute = false,
            CreateNoWindow = true
        })!;
    }

    private static Process StartShortLivedProcess()
    {
        return Process.Start(new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = "-NoProfile -Command \"Start-Sleep -Seconds 2\"",
            UseShellExecute = false,
            CreateNoWindow = true
        })!;
    }

    private static Process StartSleepingProcessWithChild(string childPidFile)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            UseShellExecute = false,
            CreateNoWindow = true
        };
        startInfo.ArgumentList.Add("-NoProfile");
        startInfo.ArgumentList.Add("-Command");
        startInfo.ArgumentList.Add(
            "$child = Start-Process powershell.exe " +
            "-ArgumentList '-NoProfile','-Command','Start-Sleep -Seconds 60' -PassThru; " +
            $"Set-Content -LiteralPath '{childPidFile.Replace("'", "''", StringComparison.Ordinal)}' -Value $child.Id; " +
            "Start-Sleep -Seconds 60");
        return Process.Start(startInfo)!;
    }

    private static async Task<Process> WaitForChildProcessAsync(string childPidFile)
    {
        var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(10);
        while (DateTime.UtcNow < deadline)
        {
            if (File.Exists(childPidFile)
                && int.TryParse(await File.ReadAllTextAsync(childPidFile), out var childProcessId))
            {
                return Process.GetProcessById(childProcessId);
            }

            await Task.Delay(100);
        }

        throw new TimeoutException("The daemon test process did not report its child PID.");
    }

    private static DaemonProcessTracker.ProcessIdentity RegisterTrackedProcesses(
        string pipeName,
        Process daemon,
        Process excel)
    {
        var daemonIdentity = DaemonProcessTracker.RegisterProcess(
            pipeName,
            daemon.Id,
            daemon.StartTime.ToUniversalTime().ToFileTimeUtc());
        DaemonProcessTracker.UpdateExcelProcesses(pipeName, daemonIdentity, [excel.Id]);
        return daemonIdentity;
    }

    private static async Task<CliResult> RunServiceStopAsync(string pipeName)
    {
        return await CliProcessHelper.RunAsync(
            "service stop",
            timeoutMs: 15000,
            environmentVariables: new Dictionary<string, string>
            {
                ["EXCELMCP_CLI_PIPE"] = pipeName
            });
    }

    private static void WriteStaleTrackingRecord(string pipeName, Process daemon, Process excel)
    {
        var trackingFile = DaemonProcessTracker.GetTrackingFilePath(pipeName);
        Directory.CreateDirectory(Path.GetDirectoryName(trackingFile)!);
        var record = new
        {
            processId = daemon.Id,
            startedAtUtcFileTime = daemon.StartTime.ToUniversalTime().ToFileTimeUtc() + 1,
            excelProcesses = new[]
            {
                new
                {
                    processId = excel.Id,
                    startedAtUtcFileTime = excel.StartTime.ToUniversalTime().ToFileTimeUtc() + 1
                }
            }
        };
        File.WriteAllText(trackingFile, JsonSerializer.Serialize(record));
    }

    private static string GetLegacyTrackingFilePath(string pipeName) =>
        Path.Combine(
            Path.GetDirectoryName(DaemonProcessTracker.GetTrackingFilePath(pipeName))!,
            $"{DaemonPipeIdentity.GetCaseSensitiveHash(pipeName)}.json");

    private static void WriteTrackingRecord(
        string trackingFile,
        Process daemon,
        Process excel)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(trackingFile)!);
        File.WriteAllText(trackingFile, JsonSerializer.Serialize(new
        {
            processId = daemon.Id,
            startedAtUtcFileTime = daemon.StartTime.ToUniversalTime().ToFileTimeUtc(),
            excelProcesses = new[]
            {
                new
                {
                    processId = excel.Id,
                    startedAtUtcFileTime = excel.StartTime.ToUniversalTime().ToFileTimeUtc()
                }
            }
        }));
    }

    private static void StopIfRunning(Process process)
    {
        if (!process.HasExited)
        {
            process.Kill(entireProcessTree: true);
            process.WaitForExit(5000);
        }
    }

    private static string GetRepositoryRoot() =>
        Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));

    private static IReadOnlyList<string> GetAdversarialPipeNames() =>
    [
        "foo",
        "FOO",
        "foo_startup",
        @"pipe/segment\with:separators?and spaces",
        $"excelmcp-{new string('x', 4096)}"
    ];

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
        var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(15);
        while (DateTime.UtcNow < deadline)
        {
            var status = await GetServiceStatusAsync(cliPath, pipeName);
            if (status.ExitCode == 0
                && status.Stdout.Contains("\"running\":true", StringComparison.Ordinal))
            {
                return;
            }

            await Task.Delay(250);
        }

        throw new TimeoutException($"Daemon for pipe '{pipeName}' did not become ready.");
    }

    private static Task<ProcessResult> GetServiceStatusAsync(string cliPath, string pipeName) =>
        RunProcessAsync(
            cliPath,
            ["service", "status", "--quiet"],
            Path.GetDirectoryName(cliPath)!,
            new Dictionary<string, string>
            {
                ["EXCELMCP_CLI_PIPE"] = pipeName
            },
            TimeSpan.FromSeconds(10));

    private static async Task StopDaemonBestEffortAsync(
        string cliPath,
        string pipeName,
        Process daemon)
    {
        if (!daemon.HasExited)
        {
            _ = await RunProcessAsync(
                cliPath,
                ["service", "stop", "--quiet"],
                Path.GetDirectoryName(cliPath)!,
                new Dictionary<string, string>
                {
                    ["EXCELMCP_CLI_PIPE"] = pipeName
                },
                TimeSpan.FromSeconds(15));
        }

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
            if (!process.HasExited)
            {
                process.Kill(entireProcessTree: true);
                await process.WaitForExitAsync();
            }

            throw new TimeoutException(
                $"Process '{fileName}' exceeded {timeout.TotalSeconds:0} seconds.");
        }

        return new ProcessResult(
            process.ExitCode,
            await stdoutTask,
            await stderrTask);
    }

    private static void CopyDirectory(string source, string destination)
    {
        Directory.CreateDirectory(destination);
        foreach (var directory in Directory.GetDirectories(
                     source,
                     "*",
                     SearchOption.AllDirectories))
        {
            Directory.CreateDirectory(
                Path.Combine(destination, Path.GetRelativePath(source, directory)));
        }

        foreach (var file in Directory.GetFiles(source, "*", SearchOption.AllDirectories))
        {
            File.Copy(
                file,
                Path.Combine(destination, Path.GetRelativePath(source, file)));
        }
    }

    private sealed record ProcessResult(int ExitCode, string Stdout, string Stderr);
}
