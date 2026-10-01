using System.Diagnostics;
using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "PluginPublication")]
public sealed class PluginPublicationScriptTests
{
    [Fact]
    public async Task PublicationAndMarketplaceScripts_PassDisposableNoWriteRegressions()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root != null && !File.Exists(Path.Combine(root.FullName, "Sbroenne.ExcelMcp.sln")))
        {
            root = root.Parent;
        }
        Assert.NotNull(root);
        // Each file owns independent disposable repositories. Bound file-level
        // concurrency so synchronous Git/PowerShell startup does not serialize the suite.
        await RunNodeAsync(root.FullName,
            [
                "--test", "--test-concurrency=3",
                "tests/ExcelMcp.SkillGeneration.Tests/PluginPublication.test.mjs",
                "tests/ExcelMcp.SkillGeneration.Tests/PluginPublicationMarketplace.test.mjs",
                "tests/ExcelMcp.SkillGeneration.Tests/PluginPublicationHistory.test.mjs",
                "tests/ExcelMcp.SkillGeneration.Tests/PluginPublicationStaging.test.mjs"
            ],
            TimeSpan.FromMinutes(3));
    }

    [Fact]
    public async Task NodeDeadline_PreservesCapturedDiagnostics()
    {
        var script = Path.Combine(Path.GetTempPath(), $"ExcelMcp.NodeDeadline.{Guid.NewGuid():N}.mjs");
        try
        {
            await File.WriteAllTextAsync(script, """
                import { test } from 'node:test';
                test('blocked fixture', () => {
                    console.log('stdout-before-timeout');
                    console.error('[plugin-publication] START blocked fixture');
                    Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 60_000);
                });
                """);
            var error = await Assert.ThrowsAsync<TimeoutException>(() => RunNodeAsync(AppContext.BaseDirectory,
                ["--test", script], TimeSpan.FromSeconds(3)));

            Assert.Contains("exceeded 3 seconds", error.Message, StringComparison.Ordinal);
            Assert.Contains("stdout-before-timeout", error.Message, StringComparison.Ordinal);
            Assert.Contains("[plugin-publication] START blocked fixture", error.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(script);
        }
    }

    [Fact]
    public async Task NodeFailure_PreservesExitCodeAndCapturedDiagnostics()
    {
        var error = await Record.ExceptionAsync(() => RunNodeAsync(AppContext.BaseDirectory,
            ["-e", "console.log('stdout-before-failure'); console.error('stderr-before-failure'); process.exitCode = 23;"],
            TimeSpan.FromSeconds(10)));

        Assert.NotNull(error);
        Assert.Contains("exit code 23", error.Message, StringComparison.Ordinal);
        Assert.Contains("stdout-before-failure", error.Message, StringComparison.Ordinal);
        Assert.Contains("stderr-before-failure", error.Message, StringComparison.Ordinal);
    }

    private static async Task RunNodeAsync(string directory, string[] arguments, TimeSpan timeout)
    {
        var info = new ProcessStartInfo("node")
        {
            WorkingDirectory = directory,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        foreach (var argument in arguments)
        {
            info.ArgumentList.Add(argument);
        }
        using var process = Process.Start(info);
        Assert.NotNull(process);
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        using var deadline = new CancellationTokenSource(timeout);
        try
        {
            await process.WaitForExitAsync(deadline.Token);
        }
        catch (OperationCanceledException) when (deadline.IsCancellationRequested)
        {
            process.Kill(true);
            await process.WaitForExitAsync();
            throw new TimeoutException(
                $"Plugin publication script tests exceeded {timeout.TotalSeconds} seconds." +
                Environment.NewLine + await stdout + Environment.NewLine + await stderr);
        }
        Assert.True(process.ExitCode == 0,
            $"Node exited with exit code {process.ExitCode}." +
            Environment.NewLine + await stdout + Environment.NewLine + await stderr);
    }
}
