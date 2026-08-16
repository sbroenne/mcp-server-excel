using System.Diagnostics;
using System.Reflection;
using Sbroenne.ExcelMcp.CLI.Commands;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Regression coverage for forced daemon shutdown reaping tracked Excel processes.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "CLI")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
public sealed class DaemonForcedStopRegressionTests
{
    [Fact]
    public async Task ForceStopTrackedDaemon_AlsoStopsTrackedExcelProcess()
    {
        var pipeName = $"excelmcp-force-stop-test-{Guid.NewGuid():N}";
        using var daemon = StartSleepingProcess();
        using var excel = StartSleepingProcess();

        try
        {
            DaemonProcessTracker.RegisterProcess(
                pipeName,
                daemon.Id,
                daemon.StartTime.ToUniversalTime().ToFileTimeUtc());
            DaemonProcessTracker.UpdateExcelProcesses(pipeName, [excel.Id]);

            var method = typeof(ServiceStopCommand).GetMethod(
                "TryForceStopTrackedDaemonAsync",
                BindingFlags.Static | BindingFlags.NonPublic);
            Assert.NotNull(method);

            var stopTask = (Task<bool>)method!.Invoke(
                null,
                [pipeName, CancellationToken.None])!;
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

        var method = typeof(ServiceStopCommand).GetMethod(
            "TryTerminateProcess",
            BindingFlags.Static | BindingFlags.NonPublic);
        Assert.NotNull(method);

        var terminated = (bool)method!.Invoke(null, [process, false])!;
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

            var exception = Record.Exception(() =>
                DaemonProcessTracker.UpdateExcelProcesses(pipeName, [Environment.ProcessId]));

            Assert.Null(exception);
            Assert.True(File.Exists(trackingFile));
        }
        finally
        {
            DaemonProcessTracker.Clear(pipeName);
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

    private static void StopIfRunning(Process process)
    {
        if (!process.HasExited)
        {
            process.Kill(entireProcessTree: true);
            process.WaitForExit(5000);
        }
    }
}
