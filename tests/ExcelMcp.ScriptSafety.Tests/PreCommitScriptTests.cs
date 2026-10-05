using System.Diagnostics;
using Sbroenne.ExcelMcp.Build;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "PreCommit")]
public sealed class PreCommitScriptTests
{
    private static readonly string RepoRoot = FindRepoRoot();

    [Theory]
    [InlineData("src/ExcelMcp.Core/Commands/Range/RangeCommands.cs", true, true)]
    [InlineData("docs/reference/report-formatting.md", true, false)]
    [InlineData("README.md", true, false)]
    [InlineData("README.md\nsrc/ExcelMcp.Core/Commands/Range/RangeCommands.cs", true, true)]
    public async Task ChangedPaths_SelectChecksWithoutCreatingPackages(string path, bool build, bool excel)
    {
        var result = await RunHookAsync(path);

        Assert.Equal(build, result.Output.Contains("dotnet build", StringComparison.Ordinal));
        Assert.Equal(excel, result.Output.Contains("excel-feature-tests-ran", StringComparison.Ordinal));
        Assert.DoesNotContain("dotnet publish", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("dotnet pack", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("npm ci", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("npm run", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("git add", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("cleanup-ran", result.Output, StringComparison.Ordinal);
        if (path == "docs/reference/report-formatting.md")
        {
            Assert.Contains("-ChangedPaths", result.Output, StringComparison.Ordinal);
        }
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Fact]
    public void ChangedPaths_ClassifyAllInputsInOneInvocation()
    {
        (string Path, bool Build, bool Excel, bool SkillTests)[] cases =
        [
            ("README.md", true, false, false),
            ("docs/guide.md", false, false, false),
            ("gh-pages/docs/index.md", false, false, false),
            ("videos/excel-mcp-intro/Capture-Evidence.ps1", true, false, false),
            ("infrastructure/azure/deploy-appinsights.ps1", true, false, false),
            ("scripts/Update-UsageAnalytics.ps1", false, false, false),
            ("vscode-extension/src/extension.ts", false, false, false),
            ("npm-packages/excelcli/package.json", false, false, false),
            ("mcpb/Build-McpBundle.ps1", true, false, false),
            (".github/plugins/excel-cli/bin/start-cli.ps1", true, false, false),
            ("scripts/Publish-PreparedPlugins.ps1", true, false, false),
            ("scripts/PluginContent.mjs", true, false, false),
            ("scripts/Update-AwesomeCopilot.mjs", true, false, false),
            (".github/workflows/update-awesome-copilot.md", true, false, false),
            (".github/workflows/publish-plugins.yml", true, false, false),
            ("scripts/Build-AgentSkills.ps1", true, false, true),
            ("tests/ExcelMcp.Core.Tests/Unit/GeneratedActionContractTests.cs", true, false, false),
            ("scripts/pre-commit.ps1", true, false, false),
            (".github/workflows/ci.yml", true, false, false),
            ("src/ExcelMcp.Core/Commands/Range/RangeCommands.cs", true, true, false),
            ("src/ExcelMcp.CLI/Program.cs", true, true, false),
            ("src/ExcelMcp.McpServer/Program.cs", true, true, false),
            ("src/ExcelMcp.Cleanup/Program.cs", true, true, false),
            ("src/ExcelMcp.Generators.Cli/Generator.cs", true, false, false),
            ("docs/reference/report-formatting.md", true, false, true),
            ("Directory.Build.props", true, true, true),
            ("Directory.Packages.props", true, true, true),
            ("global.json", true, true, true),
            (".editorconfig", true, false, true),
            ("README.md\nsrc/ExcelMcp.Core/Commands/Range/RangeCommands.cs", true, true, false),
            ("src/ExcelMcp.Core/Commands/Range/Deleted.cs\nvscode-extension/src/renamed.ts", true, true, false),
            ("docs/reference/range.md", false, false, false),
            ("skills/excel-cli-report-formatting/SKILL.md", true, false, true)
        ];
        var policy = new ValidationPolicy(RepoRoot);
        foreach (var (path, build, excel, skills) in cases)
        {
            var plan = policy.Select(path.Split('\n'));
            Assert.True(build == plan.Build, $"{path}: incorrect Build selection.");
            Assert.True(excel == plan.Excel, $"{path}: incorrect Excel selection.");
            Assert.True(skills == plan.SkillTests, $"{path}: incorrect SkillTests selection.");
            Assert.True(!excel || plan.SourceChecks, $"{path}: source guards missing.");
        }
    }

    [Theory]
    [InlineData("src/ExcelMcp.Core/Commands/Range/RangeCommands.cs")]
    [InlineData("docs/reference/report-formatting.md")]
    [InlineData("Directory.Build.props")]
    [InlineData(".editorconfig")]
    public async Task UnstagedBuildInputs_BlockBeforeBuilding(string unstaged)
    {
        var result = await RunHookAsync("src/ExcelMcp.Core/Commands/Range/RangeCommands.cs", unstaged: unstaged);
        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("differ from the index", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("dotnet build", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("git add", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public async Task UntrackedBuildInputs_BlockBeforeBuilding()
    {
        var result = await RunHookAsync("src/ExcelMcp.Core/Commands/Range/RangeCommands.cs", untracked: "src/ExcelMcp.Core/New.cs");
        Assert.NotEqual(0, result.ExitCode);
        Assert.DoesNotContain("dotnet build", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public async Task UnrelatedUnstagedDocs_DoNotBlockRuntimeChecks()
    {
        var result = await RunHookAsync("src/ExcelMcp.Core/Commands/Range/RangeCommands.cs", unstaged: "docs/guide.md");
        Assert.True(result.ExitCode == 0, result.Output);
        Assert.Contains("excel-feature-tests-ran", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("e2e-ran", result.Output, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("check-npm-lockfiles")]
    [InlineData("check-com-leaks")]
    [InlineData("check-success-flag")]
    [InlineData("check-dynamic-casts")]
    [InlineData("check-workbook-package-access")]
    [InlineData("Invoke-ExcelFreeTests")]
    [InlineData("Invoke-ExcelTests")]
    [InlineData("Test-E2E")]
    public async Task SelectedCheckFailure_IsNeverSwallowed(string failure)
    {
        var result = await RunHookAsync(failure == "Test-E2E" ? "Directory.Build.props" : "src/ExcelMcp.Core/Commands/Range/RangeCommands.cs", failure: failure);
        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("check-root-cause", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("All selected pre-commit checks passed", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public async Task BuildFailure_PreservesNativeDiagnostics()
    {
        var result = await RunHookAsync("src/ExcelMcp.Core/Commands/Range/RangeCommands.cs", failure: "dotnet");
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
        var result = await RunHookAsync("src/ExcelMcp.Core/Commands/Range/RangeCommands.cs", failure: failure);
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
            await File.WriteAllTextAsync(Path.Combine(scripts, "Invoke-BuildTool.ps1"), """
                function Invoke-ExcelMcpBuild {
                    param([string[]]$Arguments)
                    if ($Arguments[0] -ne 'build') { throw 'Unexpected hook build operation.' }
                    $file = $Arguments[[Array]::IndexOf($Arguments, '--plan') + 1]
                    $plan = Get-Content -LiteralPath $file -Raw | ConvertFrom-Json
                    $projects = if ($plan.FullSolutionBuild) { @('Sbroenne.ExcelMcp.sln') } else { @($plan.BuildProjects) }
                    foreach ($project in $projects) {
                        dotnet build $project -c Release --disable-build-servers
                        if ($LASTEXITCODE -ne 0) { throw "Selected build failed with exit code $LASTEXITCODE." }
                    }
                }
                """);
            foreach (var (name, message) in new[]
            {
                ("check-npm-lockfiles", "lockfile-ran"),
                ("check-com-leaks", "com-check-ran"),
                ("check-success-flag", "success-check-ran"),
                ("check-dynamic-casts", "casts-check-ran"),
                ("check-workbook-package-access", "package-access-check-ran"),
                ("Stop-ExcelMcpProcesses", "cleanup-ran"),
                ("Invoke-ExcelFreeTests", "excel-free-tests-ran"),
                ("Invoke-ExcelTests", "excel-feature-tests-ran"),
                ("Test-E2E", "e2e-ran"),
            })
            {
                await File.WriteAllTextAsync(Path.Combine(scripts, $"{name}.ps1"),
                    failure == "throw" && name == "check-npm-lockfiles"
                        ? "throw 'check-root-cause'"
                        : name == failure
                        ? "Write-Host 'check-root-cause'; exit 19"
                        : name == "Invoke-ExcelTests"
                        ? """
                            param([string]$PlanFile)
                            $plan = Get-Content -LiteralPath $PlanFile -Raw | ConvertFrom-Json
                            Write-Host 'excel-feature-tests-ran'
                            if ($plan.FullE2E) {
                                & (Join-Path $PSScriptRoot 'Test-E2E.ps1')
                                if ($LASTEXITCODE -ne 0) { throw "Acceptance failed with exit code $LASTEXITCODE." }
                            }
                            $global:LASTEXITCODE = 0
                            """
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
            return await RunScriptAsync(sandbox, runner);
        }
        finally { Directory.Delete(sandbox, recursive: true); }
    }

    private static async Task<(int ExitCode, string Output)> RunScriptAsync(string sandbox, string runner)
    {
        var info = new ProcessStartInfo("pwsh")
        {
            WorkingDirectory = sandbox,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
        };
        info.Environment["EXCELMCP_BUILD_ROOT"] = RepoRoot;
        info.Environment["EXCELMCP_BUILD_DLL"] = typeof(Sbroenne.ExcelMcp.Build.ValidationPolicy).Assembly.Location;
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
