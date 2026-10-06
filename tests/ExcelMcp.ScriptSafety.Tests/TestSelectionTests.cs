using System.Diagnostics;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "PreCommit")]
public sealed class TestSelectionTests
{
    [Theory]
    [InlineData("tests/ExcelMcp.SkillGeneration.Tests/Example.cs", "SkillGeneration")]
    [InlineData("tests/ExcelMcp.Packaging.Tests/Example.cs", "Packaging")]
    [InlineData("tests/ExcelMcp.ScriptSafety.Tests/Example.cs", "ScriptSafety")]
    [InlineData("scripts/Build-AgentSkills.ps1", "SkillGeneration")]
    [InlineData("docs/reference/report-formatting.md", "SkillGeneration")]
    [InlineData("scripts/Build-Plugins.ps1", "Packaging")]
    [InlineData("scripts/Update-ReleaseVersionMetadata.ps1", "Packaging")]
    [InlineData(".github/workflows/publish-mcp-registry.yml", "Packaging")]
    [InlineData("scripts/Resolve-McpRegistryRelease.ps1", "Packaging")]
    [InlineData("scripts/Test-McpRegistryPublication.ps1", "Packaging")]
    [InlineData("scripts/check-workbook-package-access.ps1", "ScriptSafety")]
    [InlineData("tests/Shared/GeneratedAssetsFixture.cs", "SkillGeneration,Packaging")]
    [InlineData("tests/Shared/PackagingScriptTestHelper.cs", "SkillGeneration,Packaging")]
    [InlineData("infrastructure/azure/deploy-appinsights.ps1", "ScriptSafety")]
    [InlineData("videos/excel-mcp-intro/Capture-Evidence.ps1", "ScriptSafety")]
    public async Task ChangedPaths_SelectOwningTestProjects(string path, string expected)
    {
        var result = await RunAsync($$"""
            . (Join-Path $root 'scripts\Get-ValidationPlan.ps1')
            $plan = Get-ValidationPlan -Paths '{{path}}'
            $projects = @(
                if ($plan.SkillTests) { 'SkillGeneration' }
                if ($plan.PackagingTests) { 'Packaging' }
                if ($plan.HookTests) { 'ScriptSafety' }
            )
            if (($projects -join ',') -ne '{{expected}}') { throw "Wrong test projects: $projects" }
            if ($plan.Excel) { throw 'Excel selected for Excel-free changes.' }
            if ($plan.FastProjects.Count -or $plan.ProcessProjects.Count -or $plan.ExcelGroups.Count) {
                throw 'Unrelated runtime tests selected.'
            }
            if ((($plan.ToolingProjects | Sort-Object) -join ',') -ne
                (('{{expected}}'.Split(',') | Sort-Object) -join ',')) {
                throw "Wrong hosted tooling projects: $($plan.ToolingProjects)"
            }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Theory]
    [InlineData("-Local -SkillTests", "SkillGeneration")]
    [InlineData("-Local -HookTests", "ScriptSafety")]
    [InlineData("-Local -PackagingTests", "Packaging")]
    [InlineData("-Local -SkillTests -HookTests -PackagingTests", "ScriptSafety,SkillGeneration,Packaging")]
    [InlineData("-Local -ChangedPaths @('tests/ExcelMcp.Packaging.Tests/Example.cs')", "Packaging")]
    [InlineData("-Local -ChangedPaths @('scripts/Build-Plugins.ps1')", "Packaging")]
    [InlineData("-Local -ChangedPaths @('scripts/Build-AgentSkills.ps1')", "SkillGeneration")]
    [InlineData("-Local -ChangedPaths @('scripts/Publish-PreparedPlugins.ps1')", "Packaging")]
    [InlineData("-Local -ChangedPaths @('tests/Shared/GeneratedAssetsFixture.cs')", "SkillGeneration,Packaging")]
    [InlineData("-Local -ChangedPaths @('tests/Shared/PackagingScriptTestHelper.cs')", "SkillGeneration,Packaging")]
    [InlineData("", "CLI,ComInterop,Core,McpServer,Service,SkillGeneration,Packaging,ScriptSafety")]
    [InlineData("-Group Tooling", "SkillGeneration,Packaging,ScriptSafety")]
    [InlineData("-Group Tooling -PlanFile $planFile", "SkillGeneration,Packaging,ScriptSafety")]
    public async Task Runner_SelectsActualProjectCommands(string arguments, string expected)
    {
        var result = await RunRunnerAsync(arguments, false);
        Assert.True(result.ExitCode == 0, result.Output);
        Assert.Contains($"selected={expected}", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Runner_SelectedFailureStopsRemainingProjects()
    {
        var result = await RunRunnerAsync("-Local -SkillTests -PackagingTests", true);
        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("SkillGeneration failed with exit code 23", result.Output, StringComparison.Ordinal);
        Assert.Contains("started=SkillGeneration", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("started=Packaging", result.Output, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("'scripts/Build-AgentSkills.ps1'", "SkillGeneration", "Feature=SkillGeneration")]
    [InlineData("'scripts/Build-Plugins.ps1'", "Packaging", "Feature=PluginBootstrap")]
    [InlineData("'scripts/check-workbook-package-access.ps1'", "ScriptSafety", "Feature=PreCommit")]
    [InlineData("'infrastructure/azure/deploy-appinsights.ps1'", "ScriptSafety", "Feature=AutomationSafety")]
    [InlineData("'doc-counts.json'", "Packaging", "FullyQualifiedName~DocumentationCounts")]
    [InlineData("'mcpb/manifest.json'", "Packaging", "Feature=McpbPackaging")]
    [InlineData("'tests/ExcelMcp.Packaging.Tests/Example.cs'", "Packaging", "RequiresExcel=false")]
    [InlineData("'tests/ExcelMcp.ScriptSafety.Tests/Example.cs'", "ScriptSafety", "RequiresExcel=false")]
    [InlineData("'tests/Shared/GeneratedAssetsFixture.cs'", "Packaging,SkillGeneration", "RequiresExcel=false")]
    [InlineData("'tests/Shared/PackagingScriptTestHelper.cs'", "Packaging,SkillGeneration", "RequiresExcel=false")]
    [InlineData("'scripts/Build-AgentSkills.ps1','scripts/check-workbook-package-access.ps1'", "ScriptSafety,SkillGeneration", "Feature=SkillGeneration")]
    [InlineData("'doc-counts.json','tests/ExcelMcp.ScriptSafety.Tests/Example.cs'", "Packaging,ScriptSafety", "FullyQualifiedName~DocumentationCounts")]
    public async Task Runner_HostedToolingSelectsOwningProjectsAndFilters(string paths, string expected, string filter)
    {
        var result = await RunRunnerAsync("-Group Tooling -PlanFile $planFile", false, paths);
        Assert.True(result.ExitCode == 0, result.Output);
        Assert.Contains($"selected={expected}", result.Output, StringComparison.Ordinal);
        Assert.Contains(filter, result.Output, StringComparison.Ordinal);
        if (paths.Contains("infrastructure/azure", StringComparison.Ordinal))
        {
            Assert.Contains("selected=ScriptSafety,ScriptSafety", result.Output, StringComparison.Ordinal);
            Assert.Contains("RunType=OnDemand&Feature=AutomationSafety", result.Output, StringComparison.Ordinal);
        }
    }

    [Fact]
    public async Task AzureInfrastructureChecks_AreSelectedOnlyForOwningChanges()
    {
        var result = await RunAsync("""
            . (Join-Path $root 'scripts\Get-ValidationPlan.ps1')
            $affected = Get-ValidationPlan -Paths @(
                'infrastructure/azure/configure-analytics-oidc.ps1',
                'infrastructure/azure/deploy-appinsights.ps1')
            if (-not $affected.AzureInfrastructureTests) { throw 'Azure safety checks were not selected.' }
            $unrelated = Get-ValidationPlan -Paths 'scripts/pre-commit.ps1'
            if ($unrelated.AzureInfrastructureTests) { throw 'Azure safety checks leaked to unrelated changes.' }
            $full = Get-ValidationPlan -Full
            if ($full.AzureInfrastructureTests) { throw 'Azure safety checks are part of normal full validation.' }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }

    [Fact]
    public async Task Runner_LocalAzureChangeRunsOnDemandChecksSeparately()
    {
        var result = await RunRunnerAsync(
            "-Local -ChangedPaths @('infrastructure/azure/configure-analytics-oidc.ps1')",
            false,
            "'infrastructure/azure/configure-analytics-oidc.ps1'");
        Assert.True(result.ExitCode == 0, result.Output);
        Assert.Contains("selected=ScriptSafety,ScriptSafety", result.Output, StringComparison.Ordinal);
        Assert.Contains("Feature=AutomationSafety", result.Output, StringComparison.Ordinal);
        Assert.Contains("RunType=OnDemand&Feature=AutomationSafety", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Runner_MixedToolingOwnersUseTheirOwnFilters()
    {
        var result = await RunRunnerAsync("-Group Tooling -PlanFile $planFile", false,
            "'doc-counts.json','tests/ExcelMcp.ScriptSafety.Tests/Example.cs'");
        Assert.True(result.ExitCode == 0, result.Output);
        Assert.Contains("Packaging : RequiresExcel=false&RunType!=OnDemand&(FullyQualifiedName~DocumentationCounts)",
            result.Output, StringComparison.Ordinal);
        Assert.Contains("ScriptSafety : RequiresExcel=false&RunType!=OnDemand&(RequiresExcel=false)",
            result.Output, StringComparison.Ordinal);
    }

    private static Task<(int ExitCode, string Output)> RunRunnerAsync(string arguments, bool fail, string? paths = null) =>
        RunAsync($$"""
            $script = Get-Content (Join-Path $root 'scripts\Invoke-ExcelFreeTests.ps1') -Raw
            $runnerDirectory = Join-Path $sandbox 'scripts'
            New-Item -ItemType Directory $runnerDirectory | Out-Null
            Copy-Item (Join-Path $root 'scripts\Get-ValidationPlan.ps1') $runnerDirectory
            $stage = Get-Content (Join-Path $root 'scripts\Invoke-TestStage.ps1') -Raw
            $boundary = '[Diagnostics.Process]::Start($info)'
            if (($stage.Split($boundary).Count - 1) -ne 1) { throw 'Process boundary changed.' }
            Set-Content (Join-Path $runnerDirectory 'Invoke-TestStage.ps1') $stage.Replace($boundary, '(Start-TestProcess $info)')
            $runner = Join-Path $runnerDirectory 'Invoke-ExcelFreeTests.ps1'
            Set-Content $runner $script
            . (Join-Path $runnerDirectory 'Get-ValidationPlan.ps1')
            $planFile = Join-Path $sandbox 'plan.json'
            Get-ValidationPlan {{(paths is null ? "-Full" : $"-Paths @({paths})")}} | ConvertTo-Json -Depth 10 | Set-Content $planFile
            $global:selected = [Collections.Generic.List[string]]::new()
            function Start-TestProcess($info) {
                $arguments = @($info.ArgumentList)
                if ($info.FileName -ne 'dotnet' -or $arguments[0] -ne 'test') { throw 'Unexpected command.' }
                $project = [regex]::Match($arguments[1], 'ExcelMcp\.(\w+)\.Tests\.csproj$').Groups[1].Value
                if (-not $project) { throw 'Invalid project path.' }
                $global:selected.Add($project)
                Write-Host "started=$project"
                $filter = $arguments[[Array]::IndexOf($arguments, '--filter') + 1]
                if ($filter -notmatch 'RequiresExcel=false&' -or
                    ($filter -notmatch 'RunType!=OnDemand' -and $filter -notmatch 'RunType=OnDemand')) {
                    throw 'Classification filter lost.'
                }
                if ($filter -match 'RunType=OnDemand' -and $filter -notmatch 'Feature=AutomationSafety') {
                    throw 'Unexpected on-demand tooling selection.'
                }
                $results = $arguments[[Array]::IndexOf($arguments, '--results-directory') + 1]
                New-Item -ItemType Directory $results -Force | Out-Null
                $logger = $arguments[[Array]::IndexOf($arguments, '--logger') + 1]
                $report = [regex]::Match($logger, 'LogFileName=([^;]+)').Groups[1].Value
                if (-not $report) { throw 'Missing report name.' }
                Set-Content (Join-Path $results $report) '<TestRun><ResultSummary outcome="Completed"><Counters total="1" passed="1" /></ResultSummary></TestRun>'
                $process = [pscustomobject]@{
                    ExitCode = {{(fail ? 23 : 0)}}
                    StandardOutput = [IO.StringReader]::new('')
                    StandardError = [IO.StringReader]::new('')
                }
                $process | Add-Member ScriptMethod WaitForExit { param($timeout) return $true }
                $process | Add-Member ScriptMethod Dispose {}
                return $process
            }
            & $runner {{arguments}}
            Write-Output "selected=$($global:selected -join ',')"
            """);

    private static async Task<(int ExitCode, string Output)> RunAsync(string body)
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root != null && !File.Exists(Path.Combine(root.FullName, "Sbroenne.ExcelMcp.sln")))
        {
            root = root.Parent;
        }
        Assert.NotNull(root);
        var sandbox = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), $"ExcelMcp.Selection.{Guid.NewGuid():N}")).FullName;
        try
        {
            var runner = Path.Combine(sandbox, "test.ps1");
            await File.WriteAllTextAsync(runner, $"""
                $ErrorActionPreference = 'Stop'
                $root = '{root.FullName.Replace("'", "''", StringComparison.Ordinal)}'
                $sandbox = '{sandbox.Replace("'", "''", StringComparison.Ordinal)}'
                {body}
                """);
            var info = new ProcessStartInfo("pwsh")
            {
                WorkingDirectory = sandbox,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false
            };
            foreach (var argument in new[] { "-NoProfile", "-File", runner }) { info.ArgumentList.Add(argument); }
            using var process = Process.Start(info)!;
            var stdout = process.StandardOutput.ReadToEndAsync();
            var stderr = process.StandardError.ReadToEndAsync();
            using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            try { await process.WaitForExitAsync(deadline.Token); }
            catch (OperationCanceledException)
            {
                process.Kill(true);
                await process.WaitForExitAsync();
                throw new TimeoutException("Test selection exceeded 30 seconds.");
            }
            return (process.ExitCode, await stdout + await stderr);
        }
        finally { Directory.Delete(sandbox, true); }
    }
}
