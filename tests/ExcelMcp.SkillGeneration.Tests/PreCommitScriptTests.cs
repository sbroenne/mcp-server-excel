using System.Diagnostics;
using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

public sealed class PreCommitScriptTests
{
    private static readonly string RepoRoot = FindRepoRoot();

    [Theory]
    [InlineData("FEATURES.md", false)]
    [InlineData("gh-pages/hooks.py", false)]
    [InlineData("scripts/pre-commit.ps1", false)]
    [InlineData("scripts/check-doc-counts.ps1", true)]
    [InlineData("tests/ExcelMcp.SkillGeneration.Tests/PreCommitScriptTests.cs", true)]
    [InlineData(".github/workflows/ci.yml", true)]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PreCommit")]
    public async Task ValidationOnlyChanges_DoNotPackageReleases(string path, bool requiresBuild)
    {
        var result = await RunHookAsync(path);

        Assert.True(result.ExitCode == 0, result.CombinedOutput);
        Assert.DoesNotContain("Building CLI release deliverables", result.CombinedOutput, StringComparison.Ordinal);
        Assert.Equal(requiresBuild, result.CombinedOutput.Contains("Building Release solution", StringComparison.Ordinal));
        Assert.Equal(requiresBuild, result.CombinedOutput.Contains("Validating documentation tool/operation counts", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("src/ExcelMcp.CLI/Program.cs", true)]
    [InlineData("src/ExcelMcp.Service/ExcelMcpService.cs", true)]
    [InlineData("src/ExcelMcp.Generators/Generator.cs", true)]
    [InlineData("Directory.Build.props", false)]
    [InlineData("scripts/Build-AgentSkills.ps1", false)]
    [InlineData(".github/workflows/release.yml", false)]
    [InlineData("unknown-build-input.json", false)]
    [InlineData("scripts/check-doc-counts.ps1\nsrc/ExcelMcp.Core/Command.cs", true)]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PreCommit")]
    public async Task ShippingChanges_StillPackageAndReportPublishFailure(string paths, bool requiresExcel)
    {
        var result = await RunHookAsync(paths);

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("Building CLI release deliverables", result.CombinedOutput, StringComparison.Ordinal);
        Assert.Contains("publish-root-cause", result.CombinedOutput, StringComparison.Ordinal);
        Assert.Contains("exit code 23", result.CombinedOutput, StringComparison.Ordinal);
        Assert.DoesNotContain("Cannot find path", result.CombinedOutput, StringComparison.Ordinal);
        Assert.Equal(requiresExcel, result.CombinedOutput.Contains("Running Excel-dependent E2E tests", StringComparison.Ordinal));
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PreCommit")]
    public async Task PackagingException_RetainsEarlierCommandOutput()
    {
        var result = await RunHookAsync("Directory.Build.props", failureExitCode: 0, createArtifacts: false);

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("publish-root-cause", result.CombinedOutput, StringComparison.Ordinal);
        Assert.Contains("excelcli.exe", result.CombinedOutput, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("pack", "CLI", "CLI")]
    [InlineData("publish", "CLI", "CLI")]
    [InlineData("pack", "McpServer", "MCP Server")]
    [InlineData("publish", "McpServer", "MCP Server")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PreCommit")]
    public async Task NativePackagingFailure_ReportsCommandAndCapturedOutput(string command, string project, string label)
    {
        var result = await RunHookAsync("Directory.Build.props", failureCommand: command, failureProject: project);

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains($"dotnet {command} ({label}) failed with exit code 23", result.CombinedOutput, StringComparison.Ordinal);
        Assert.Contains($"{command}-root-cause", result.CombinedOutput, StringComparison.Ordinal);
        Assert.Contains($"{command}-stderr-cause", result.CombinedOutput, StringComparison.Ordinal);
        Assert.DoesNotContain("Cannot find path", result.CombinedOutput, StringComparison.Ordinal);
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PreCommit")]
    public async Task MergeCommit_SelectsChangesAgainstMergeParent()
    {
        var result = await RunHookAsync("FEATURES.md", merging: true);

        Assert.True(result.ExitCode == 0, result.CombinedOutput);
        Assert.Contains("comparison-base=merge-parent", result.CombinedOutput, StringComparison.Ordinal);
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PreCommit")]
    public async Task NpmLockfileFailure_BlocksEvenDocumentationOnlyCommitsBeforeCleanup()
    {
        var result = await RunHookAsync("FEATURES.md", npmLockfileFailure: true);

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("npm-staged=True", result.CombinedOutput, StringComparison.Ordinal);
        Assert.Contains("Npm lockfiles contain fixed download URLs", result.CombinedOutput, StringComparison.Ordinal);
        Assert.DoesNotContain("Stopping pipe-owned", result.CombinedOutput, StringComparison.Ordinal);
    }

    private static async Task<ScriptResult> RunHookAsync(
        string paths, int failureExitCode = 23, bool merging = false,
        string failureCommand = "publish", string failureProject = "CLI", bool createArtifacts = true,
        bool npmLockfileFailure = false)
    {
        var sandbox = Path.Combine(Path.GetTempPath(), $"ExcelMcpPreCommit-{Guid.NewGuid():N}");
        var scripts = Path.Combine(sandbox, "scripts");
        Directory.CreateDirectory(scripts);
        try
        {
            File.Copy(Path.Combine(RepoRoot, "scripts", "pre-commit.ps1"), Path.Combine(scripts, "pre-commit.ps1"));
            await File.WriteAllTextAsync(Path.Combine(sandbox, "Directory.Build.props"),
                "<Project><PropertyGroup><Version>1.0.0</Version></PropertyGroup></Project>");
            await File.WriteAllTextAsync(Path.Combine(sandbox, "README.md"), "test readme");
            await File.WriteAllTextAsync(Path.Combine(sandbox, "LICENSE"), "test license");
            foreach (var name in new[]
            {
                "Stop-ExcelMcpProcesses", "check-com-leaks", "audit-core-coverage",
                "check-mcp-core-implementations", "check-success-flag", "Build-BootstrapScripts",
                "check-doc-counts", "Test-E2E", "check-plugin-readmes", "check-dynamic-casts"
            })
            {
                await File.WriteAllTextAsync(Path.Combine(scripts, $"{name}.ps1"), "$global:LASTEXITCODE = 0");
            }
            await File.WriteAllTextAsync(Path.Combine(scripts, "check-npm-lockfiles.ps1"), $$"""
                param([switch]$Staged)
                Write-Output "npm-staged=$Staged"
                $global:LASTEXITCODE = {{(npmLockfileFailure ? 1 : 0)}}
                """);

            // Exercise the real hook in isolation: no builds, Excel processes, or real Git state.
            var harness = $$"""
                function global:git {
                    $global:LASTEXITCODE = 0
                    switch ($args[0]) {
                        'rev-parse' {
                            if ({{(merging ? "$true" : "$false")}}) { 'merge-parent' }
                            else { $global:LASTEXITCODE = 1 }
                        }
                        'branch' { 'test-branch' }
                        'diff' {
                            if ($args -contains '--cached') {
                                [Console]::WriteLine("comparison-base=$($args[-1])")
                                '{{paths.Replace("'", "''", StringComparison.Ordinal)}}' -split "`n"
                            }
                        }
                    }
                }
                function global:dotnet {
                    $global:LASTEXITCODE = 0
                    if ($args[0] -eq 'build' -and
                        ($args -contains '--configfile' -or $args -contains '--source')) {
                        throw 'Build must preserve inherited package sources.'
                    }
                    if ($args[0] -eq '{{failureCommand}}' -and $args[1] -like '*ExcelMcp.{{failureProject}}\*') {
                        Write-Output '{{failureCommand}}-root-cause'
                        Write-Error '{{failureCommand}}-stderr-cause' -ErrorAction Continue
                        $global:LASTEXITCODE = {{failureExitCode}}
                        return
                    }
                    if ({{(createArtifacts ? "$true" : "$false")}} -and $args[0] -in @('pack', 'publish')) {
                        $output = $args[[Array]::IndexOf($args, '--output') + 1]
                        $name = if ($args[0] -eq 'pack') { 'test.nupkg' }
                            elseif ($args[1] -like '*ExcelMcp.CLI\*') { 'excelcli.exe' }
                            else { 'Sbroenne.ExcelMcp.McpServer.exe' }
                        [IO.File]::WriteAllText((Join-Path $output $name), 'test artifact')
                    }
                }
                & '.\scripts\pre-commit.ps1'
                """;

            var startInfo = new ProcessStartInfo
            {
                FileName = "pwsh",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                WorkingDirectory = sandbox
            };
            startInfo.ArgumentList.Add("-NoProfile");
            startInfo.ArgumentList.Add("-Command");
            startInfo.ArgumentList.Add(harness);
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

            return new ScriptResult(process.ExitCode, $"{await stdout}{Environment.NewLine}{await stderr}");
        }
        finally
        {
            Directory.Delete(sandbox, recursive: true);
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
        throw new DirectoryNotFoundException("Could not locate repository root from test output directory.");
    }

    private sealed record ScriptResult(int ExitCode, string CombinedOutput);
}
