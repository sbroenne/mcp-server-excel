using System.Diagnostics;
using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

[Trait("Layer", "Build")]
[Trait("Category", "Unit")]
[Trait("Feature", "ReleaseMetadata")]
public sealed class McpbStagingCleanupTests
{
    private static readonly string RepoRoot = FindRepoRoot();
    private static readonly string BuildScript = Path.Combine(
        RepoRoot,
        "mcpb",
        "Build-McpBundle.ps1");

    [Fact]
    public async Task RemoveStagingDirectory_TransientLock_RechecksUntilRemoved()
    {
        var stagingDirectory = CreateStagingDirectory();
        try
        {
            var result = await RunCleanupScenarioAsync(
                stagingDirectory,
                """
                $script:attempt = 0
                $deleteAction = {
                    param($targetPath)
                    $script:attempt++
                    if ($script:attempt -lt 3) {
                        throw [System.IO.IOException]::new('scanner lock')
                    }
                    [System.IO.Directory]::Delete($targetPath, $true)
                }
                $delayAction = { param($milliseconds) }
                Remove-StagingDirectory -Path $stagingPath -Timeout ([TimeSpan]::FromSeconds(5)) -DeleteAction $deleteAction -DelayAction $delayAction
                Write-Output "attempts=$script:attempt"
                """);

            Assert.True(result.ExitCode == 0, result.CombinedOutput);
            Assert.Contains("attempts=3", result.Stdout, StringComparison.Ordinal);
            Assert.False(Directory.Exists(stagingDirectory));
        }
        finally
        {
            if (Directory.Exists(stagingDirectory))
            {
                Directory.Delete(stagingDirectory, recursive: true);
            }
        }
    }

    [Fact]
    public async Task RemoveStagingDirectory_PersistentLock_FailsAtDeadlineWithDiagnostics()
    {
        var stagingDirectory = CreateStagingDirectory();
        try
        {
            var result = await RunCleanupScenarioAsync(
                stagingDirectory,
                """
                $script:now = [DateTime]::Parse('2026-01-01T00:00:00Z').ToUniversalTime()
                $deleteAction = {
                    param($targetPath)
                    throw [System.IO.IOException]::new('scanner lock persisted')
                }
                $delayAction = {
                    param($milliseconds)
                    $script:now = $script:now.AddMilliseconds($milliseconds)
                }
                $utcNowAction = { $script:now }
                try {
                    Remove-StagingDirectory -Path $stagingPath -Timeout ([TimeSpan]::FromSeconds(1)) -DeleteAction $deleteAction -DelayAction $delayAction -UtcNowAction $utcNowAction
                }
                catch {
                    [Console]::Error.WriteLine($_.Exception.Message)
                    exit 23
                }
                """);

            Assert.Equal(23, result.ExitCode);
            Assert.Contains(stagingDirectory, result.Stderr, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("within 1 seconds", result.Stderr, StringComparison.Ordinal);
            Assert.Contains("stale staging remains", result.Stderr, StringComparison.Ordinal);
            Assert.Contains("scanner lock persisted", result.Stderr, StringComparison.Ordinal);
            Assert.True(Directory.Exists(stagingDirectory));
        }
        finally
        {
            if (Directory.Exists(stagingDirectory))
            {
                Directory.Delete(stagingDirectory, recursive: true);
            }
        }
    }

    private static string CreateStagingDirectory()
    {
        var path = Path.Combine(
            Path.GetTempPath(),
            $"excelmcp-mcpb-cleanup-{Guid.NewGuid():N}");
        Directory.CreateDirectory(path);
        File.WriteAllText(Path.Combine(path, "excel-mcp-server.exe"), "test");
        return path;
    }

    private static async Task<ScriptResult> RunCleanupScenarioAsync(
        string stagingDirectory,
        string scenario)
    {
        var testScript = Path.Combine(
            Path.GetTempPath(),
            $"excelmcp-mcpb-cleanup-{Guid.NewGuid():N}.ps1");
        var escapedBuildScript = BuildScript.Replace("'", "''", StringComparison.Ordinal);
        var escapedStagingDirectory = stagingDirectory.Replace("'", "''", StringComparison.Ordinal);
        await File.WriteAllTextAsync(
            testScript,
            $"""
            $ErrorActionPreference = 'Stop'
            . '{escapedBuildScript}'
            $stagingPath = '{escapedStagingDirectory}'
            {scenario}
            """);

        try
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = "pwsh",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                WorkingDirectory = RepoRoot
            };
            startInfo.ArgumentList.Add("-NoProfile");
            startInfo.ArgumentList.Add("-ExecutionPolicy");
            startInfo.ArgumentList.Add("Bypass");
            startInfo.ArgumentList.Add("-File");
            startInfo.ArgumentList.Add(testScript);

            using var process = Process.Start(startInfo);
            Assert.NotNull(process);
            var stdout = process.StandardOutput.ReadToEndAsync();
            var stderr = process.StandardError.ReadToEndAsync();
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            await process.WaitForExitAsync(timeout.Token);
            return new ScriptResult(process.ExitCode, await stdout, await stderr);
        }
        finally
        {
            File.Delete(testScript);
        }
    }

    private static string FindRepoRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Sbroenne.ExcelMcp.sln")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException(
            "Could not locate repository root from test output directory.");
    }

    private sealed record ScriptResult(int ExitCode, string Stdout, string Stderr)
    {
        public string CombinedOutput => $"{Stdout}{Environment.NewLine}{Stderr}";
    }
}
