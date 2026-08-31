using System.Diagnostics;
using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

/// <summary>
/// Integration tests for MCPB packaging script behavior.
/// </summary>
public sealed class McpbPackagingScriptTests
{
    private static readonly string RepoRoot = FindRepoRoot();
    private static readonly string PackagingHelpers = Path.Combine(
        RepoRoot,
        "mcpb",
        "McpbPackaging.ps1");

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "McpbPackaging")]
    public async Task RemoveStagingDirectory_RetriesTransientFileLockUntilDirectoryIsGone()
    {
        var sandbox = CreateSandbox();
        try
        {
            var script = $$"""
                $ErrorActionPreference = 'Stop'
                . '{{EscapePowerShellLiteral(PackagingHelpers)}}'
                $script:attempts = 0
                $removeDirectory = {
                    param([string]$Path)
                    $script:attempts++
                    if ($script:attempts -lt 3) {
                        throw [System.IO.IOException]::new('locked by scanner')
                    }

                    [System.IO.Directory]::Delete($Path, $true)
                }

                Remove-McpbStagingDirectory `
                    -Path '{{EscapePowerShellLiteral(sandbox)}}' `
                    -Timeout ([TimeSpan]::FromSeconds(1)) `
                    -RetryInterval ([TimeSpan]::Zero) `
                    -RemoveDirectory $removeDirectory
                Write-Output "attempts=$script:attempts"
                """;

            var result = await RunPowerShellAsync(script);

            Assert.True(result.ExitCode == 0, result.CombinedOutput);
            Assert.Contains("attempts=3", result.Stdout, StringComparison.Ordinal);
            Assert.False(Directory.Exists(sandbox));
        }
        finally
        {
            if (Directory.Exists(sandbox))
            {
                Directory.Delete(sandbox, recursive: true);
            }
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "McpbPackaging")]
    public async Task RemoveStagingDirectory_WhenTimeoutExpires_ReportsTerminalState()
    {
        var sandbox = CreateSandbox();
        try
        {
            var script = $$"""
                $ErrorActionPreference = 'Stop'
                . '{{EscapePowerShellLiteral(PackagingHelpers)}}'
                $removeDirectory = {
                    param([string]$Path)
                    throw [System.UnauthorizedAccessException]::new('locked by scanner')
                }

                Remove-McpbStagingDirectory `
                    -Path '{{EscapePowerShellLiteral(sandbox)}}' `
                    -Timeout ([TimeSpan]::Zero) `
                    -RetryInterval ([TimeSpan]::Zero) `
                    -RemoveDirectory $removeDirectory
                """;

            var result = await RunPowerShellAsync(script);

            Assert.NotEqual(0, result.ExitCode);
            Assert.True(Directory.Exists(sandbox));
            Assert.Contains(
                "Failed to remove MCPB staging directory",
                result.Stderr,
                StringComparison.Ordinal);
            Assert.Contains(sandbox, result.Stderr, StringComparison.Ordinal);
            Assert.Contains("after 1 attempt", result.Stderr, StringComparison.Ordinal);
            Assert.Contains("within 0 ms", result.Stderr, StringComparison.Ordinal);
            Assert.Contains("stale staging remains", result.Stderr, StringComparison.Ordinal);
            Assert.Contains("locked by scanner", result.Stderr, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(sandbox, recursive: true);
        }
    }

    private static string CreateSandbox()
    {
        var sandbox = Path.Combine(Path.GetTempPath(), $"ExcelMcpMcpbPackaging-{Guid.NewGuid():N}");
        Directory.CreateDirectory(Path.Combine(sandbox, "server"));
        File.WriteAllText(Path.Combine(sandbox, "server", "excel-mcp-server.exe"), "test");
        return sandbox;
    }

    private static string EscapePowerShellLiteral(string value) =>
        value.Replace("'", "''", StringComparison.Ordinal);

    private static async Task<ScriptResult> RunPowerShellAsync(string script)
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
        startInfo.ArgumentList.Add("-Command");
        startInfo.ArgumentList.Add(script);

        using var process = Process.Start(startInfo);
        Assert.NotNull(process);

        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        try
        {
            await process.WaitForExitAsync(timeout.Token);
        }
        catch (OperationCanceledException)
        {
            process.Kill(entireProcessTree: true);
            await process.WaitForExitAsync();
            throw;
        }

        return new ScriptResult(process.ExitCode, await stdout, await stderr);
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

        throw new DirectoryNotFoundException("Could not locate repository root from test output directory.");
    }

    private sealed record ScriptResult(int ExitCode, string Stdout, string Stderr)
    {
        public string CombinedOutput => $"{Stdout}{Environment.NewLine}{Stderr}";
    }
}
