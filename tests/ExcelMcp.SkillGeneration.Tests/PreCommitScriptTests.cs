using System.Diagnostics;
using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "PreCommit")]
public sealed class PreCommitScriptTests
{
    private static readonly string RepoRoot = FindRepoRoot();

    [Theory]
    [InlineData("README.md", false, false)]
    [InlineData("docs/guide.md", false, false)]
    [InlineData("gh-pages/docs/index.md", false, false)]
    [InlineData("videos/excel-mcp-intro/Capture-Evidence.ps1", false, false)]
    [InlineData("infrastructure/azure/deploy-appinsights.ps1", false, false)]
    [InlineData("scripts/Update-UsageAnalytics.ps1", false, false)]
    [InlineData("vscode-extension/src/extension.ts", false, false)]
    [InlineData("npm-packages/excelcli/package.json", false, false)]
    [InlineData("mcpb/Build-McpBundle.ps1", false, false)]
    [InlineData(".github/plugins/_shared/download.ps1.template", false, false)]
    [InlineData("scripts/Build-AgentSkills.ps1", true, false)]
    [InlineData("tests/ExcelMcp.Core.Tests/ExampleTests.cs", true, false)]
    [InlineData("scripts/pre-commit.ps1", true, false)]
    [InlineData(".github/workflows/ci.yml", true, false)]
    [InlineData("src/ExcelMcp.Core/Command.cs", true, true)]
    [InlineData("src/ExcelMcp.CLI/Program.cs", true, true)]
    [InlineData("src/ExcelMcp.Cleanup/Program.cs", true, true)]
    [InlineData("src/ExcelMcp.Generators.Cli/Generator.cs", true, true)]
    [InlineData("Directory.Build.props", true, true)]
    [InlineData("Directory.Packages.props", true, true)]
    [InlineData("global.json", true, true)]
    [InlineData(".editorconfig", true, false)]
    [InlineData("README.md\nsrc/ExcelMcp.Core/Command.cs", true, true)]
    [InlineData("src/ExcelMcp.Core/Deleted.cs\nvscode-extension/src/renamed.ts", true, true)]
    [InlineData("skills/shared/range.md", true, true)]
    [InlineData("unknown-build-input.config", true, true)]
    public async Task ChangedPaths_SelectChecksWithoutCreatingPackages(string path, bool build, bool excel)
    {
        var result = await RunHookAsync(path);

        Assert.Equal(build, result.Output.Contains("dotnet build", StringComparison.Ordinal));
        Assert.Equal(excel, result.Output.Contains("e2e-ran", StringComparison.Ordinal));
        Assert.DoesNotContain("dotnet publish", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("dotnet pack", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("npm ci", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("npm run", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("git add", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("cleanup-ran", result.Output, StringComparison.Ordinal);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Fact]
    public async Task RuntimeChange_NeverCreatesDistributablePackages()
    {
        var result = await RunHookAsync("src/ExcelMcp.Core/Command.cs");
        Assert.True(result.ExitCode == 0, result.Output);
        Assert.DoesNotContain("dotnet publish", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("dotnet pack", result.Output, StringComparison.Ordinal);
        Assert.Contains("non-packaging-tests-ran", result.Output, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("src/ExcelMcp.Core/Command.cs")]
    [InlineData("skills/shared/range.md")]
    [InlineData("Directory.Build.props")]
    [InlineData(".editorconfig")]
    public async Task UnstagedBuildInputs_BlockBeforeBuilding(string unstaged)
    {
        var result = await RunHookAsync("src/ExcelMcp.Core/Command.cs", unstaged: unstaged);
        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("differ from the index", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("dotnet build", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("git add", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public async Task UntrackedBuildInputs_BlockBeforeBuilding()
    {
        var result = await RunHookAsync("src/ExcelMcp.Core/Command.cs", untracked: "src/ExcelMcp.Core/New.cs");
        Assert.NotEqual(0, result.ExitCode);
        Assert.DoesNotContain("dotnet build", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public async Task UnrelatedUnstagedDocs_DoNotBlockRuntimeChecks()
    {
        var result = await RunHookAsync("src/ExcelMcp.Core/Command.cs", unstaged: "docs/guide.md");
        Assert.True(result.ExitCode == 0, result.Output);
        Assert.Contains("e2e-ran", result.Output, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("check-npm-lockfiles")]
    [InlineData("check-com-leaks")]
    [InlineData("check-success-flag")]
    [InlineData("check-dynamic-casts")]
    [InlineData("check-workbook-package-access")]
    [InlineData("Invoke-ExcelFreeTests")]
    [InlineData("Test-E2E")]
    public async Task SelectedCheckFailure_IsNeverSwallowed(string failure)
    {
        var result = await RunHookAsync("src/ExcelMcp.Core/Command.cs", failure: failure);
        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("check-root-cause", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("All selected pre-commit checks passed", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public async Task BuildFailure_PreservesNativeDiagnostics()
    {
        var result = await RunHookAsync("src/ExcelMcp.Core/Command.cs", failure: "dotnet");
        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("build-root-cause", result.Output, StringComparison.Ordinal);
        Assert.Contains("exit code 23", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public async Task MergeCommit_UsesMergeParentAndStillGuardsLockfiles()
    {
        var result = await RunHookAsync("README.md", merge: true);
        Assert.True(result.ExitCode == 0, result.Output);
        Assert.Contains("comparison-base=merge-parent", result.Output, StringComparison.Ordinal);
        Assert.Contains("lockfile-ran -Staged", result.Output, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("git")]
    [InlineData("throw")]
    public async Task InspectionOrCheckException_StopsValidation(string failure)
    {
        var result = await RunHookAsync("src/ExcelMcp.Core/Command.cs", failure: failure);
        Assert.NotEqual(0, result.ExitCode);
        Assert.DoesNotContain("All selected pre-commit checks passed", result.Output, StringComparison.Ordinal);
        Assert.Contains("root-cause", result.Output, StringComparison.Ordinal);
    }

    private static async Task<(int ExitCode, string Output)> RunHookAsync(
        string staged, string unstaged = "", string untracked = "", string failure = "", bool merge = false)
    {
        var sandbox = Path.Combine(Path.GetTempPath(), $"ExcelMcp.Hook.{Guid.NewGuid():N}");
        var scripts = Path.Combine(sandbox, "scripts");
        Directory.CreateDirectory(scripts);
        try
        {
            foreach (var name in new[] { "pre-commit", "Get-ValidationPlan" })
            {
                File.Copy(Path.Combine(RepoRoot, "scripts", $"{name}.ps1"), Path.Combine(scripts, $"{name}.ps1"));
            }
            foreach (var (name, message) in new[]
            {
                ("check-npm-lockfiles", "lockfile-ran"),
                ("check-com-leaks", "com-check-ran"),
                ("check-success-flag", "success-check-ran"),
                ("check-dynamic-casts", "casts-check-ran"),
                ("check-workbook-package-access", "package-access-check-ran"),
                ("Stop-ExcelMcpProcesses", "cleanup-ran"),
                ("Invoke-ExcelFreeTests", "non-packaging-tests-ran"),
                ("Test-E2E", "e2e-ran"),
            })
            {
                await File.WriteAllTextAsync(Path.Combine(scripts, $"{name}.ps1"),
                    failure == "throw" && name == "check-npm-lockfiles"
                        ? "throw 'check-root-cause'"
                        : name == failure
                        ? "Write-Host 'check-root-cause'; exit 19"
                        : $"Write-Host \"{message} $args\"; $global:LASTEXITCODE = 0");
            }
            var runner = Path.Combine(sandbox, "run.ps1");
            await File.WriteAllTextAsync(runner, $$"""
                $ErrorActionPreference = 'Stop'
                function git {
                    $global:LASTEXITCODE = 0
                    Write-Host "git $args"
                    if ($args[0] -eq 'branch') { return 'test-branch' }
                    if ($args[0] -eq 'rev-parse') {
                        if ({{(merge ? "$true" : "$false")}}) { return 'merge-parent' }
                        $global:LASTEXITCODE = 1; return
                    }
                    if ($args -contains '--cached') {
                        Write-Host "comparison-base=$($args[-1])"
                        if ('{{failure}}' -eq 'git') { Write-Host 'git-root-cause'; $global:LASTEXITCODE=23; return }
                        return '{{staged.Replace("'", "''", StringComparison.Ordinal)}}' -split '\r?\n'
                    }
                    if ($args -contains 'diff') { return '{{unstaged.Replace("'", "''", StringComparison.Ordinal)}}' }
                    if ($args -contains 'ls-files') { return '{{untracked.Replace("'", "''", StringComparison.Ordinal)}}' }
                    throw "Unexpected git command: $args"
                }
                function dotnet {
                    Write-Host "dotnet $args"
                    if ('{{failure}}' -eq 'dotnet') {
                        Write-Host 'build-root-cause'; $global:LASTEXITCODE = 23; return
                    }
                    $global:LASTEXITCODE = 0
                }
                function npm { throw "Packaging must not run: npm $args" }
                & (Join-Path $PSScriptRoot 'scripts\pre-commit.ps1')
                exit $LASTEXITCODE
                """);
            var info = new ProcessStartInfo("pwsh")
            {
                WorkingDirectory = sandbox,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
            };
            foreach (var argument in new[] { "-NoLogo", "-NoProfile", "-File", runner }) { info.ArgumentList.Add(argument); }
            using var process = Process.Start(info)!;
            var stdout = process.StandardOutput.ReadToEndAsync();
            var stderr = process.StandardError.ReadToEndAsync();
            using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            try { await process.WaitForExitAsync(deadline.Token); }
            catch (OperationCanceledException)
            {
                process.Kill(entireProcessTree: true);
                await process.WaitForExitAsync();
                throw new TimeoutException("Pre-commit regression exceeded 30 seconds.");
            }
            return (process.ExitCode, await stdout + await stderr);
        }
        finally { Directory.Delete(sandbox, recursive: true); }
    }

    private static string FindRepoRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Sbroenne.ExcelMcp.sln"))) { return directory.FullName; }
            directory = directory.Parent;
        }
        throw new InvalidOperationException("Cannot find repository root.");
    }
}
