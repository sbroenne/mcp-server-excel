using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class TestSelectionTests
{
    [Fact]
    public async Task Commit_ExcludesSystemSlowToolingAndInfrastructureTests()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            . .\scripts\ValidationExecution.ps1
            $plan = Get-ValidationPlan -Full
            $selected = @(Get-FreeTestSelection -Plan $plan -Group Commit)
            if (($selected.Owner -join ',') -cne 'CLI,Core,McpServer,Service') { throw 'Commit owners changed.' }
            foreach ($item in $selected) {
                foreach ($required in @('RequiresExcel=false','RunType!=OnDemand','AdapterTestKind!=System','Speed!=Slow')) {
                    if (-not $item.Filter.Contains($required)) { throw "Commit exclusion missing: $required" }
                }
            }
            $ci = @(Get-FreeTestSelection -Plan $plan -Group Fast)
            if (@($ci | Where-Object { $_.Filter.Contains('Speed!=Slow') }).Count) { throw 'CI lost slow cases.' }
            $process = @(Get-FreeTestSelection -Plan $plan -Group Process)
            if ($process.Count -ne 1 -or $process[0].Owner -ne 'CLI' -or
                -not $process[0].Filter.Contains('AdapterTestKind=System')) { throw 'Process partition lost.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task ToolingChanges_SelectTheirOwnersWithoutWorkbookTests()
    {
        var run = await ValidationSelectionTests.RunCasesAsync(
        [
            ["scripts/Build-AgentSkills.ps1", "Packaging,SkillGeneration"],
            ["docs/reference/report-formatting.md", "Packaging,SkillGeneration"],
            ["scripts/Build-Plugins.ps1", "Packaging"],
            ["scripts/Update-ReleaseVersionMetadata.ps1", "Packaging"],
            [".github/workflows/publish-mcp-registry.yml", "Packaging"],
            ["mcpb/README.md", "Packaging"],
            ["src/ExcelMcp.Analyzers/ComAccessAnalyzer.cs", "ScriptSafety"],
            ["tests/Shared/GeneratedAssetsFixture.cs", "Packaging,SkillGeneration"],
            ["tests/Shared/PackagingScriptTestHelper.cs", "Packaging,SkillGeneration"],
            ["infrastructure/azure/deploy-appinsights.ps1", "ScriptSafety"],
            ["videos/excel-mcp-intro/Capture-Evidence.ps1", "ScriptSafety"]
        ], """
            $path, $owners = $case
            . .\scripts\ValidationExecution.ps1
            $plan = Get-ValidationPlan -Paths $path
            $selected = @(Get-FreeTestSelection -Plan $plan -Group Tooling)
            if (($selected.Owner -join ',') -cne $owners -or $plan.Excel -or
                $plan.FastProjects.Count -or $plan.ProcessProjects.Count) { throw 'Wrong tooling consumers.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task MixedOwners_KeepIndependentFiltersAndStopAfterFailure()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            . .\scripts\ValidationExecution.ps1
            $plan = Get-ValidationPlan -Paths @('scripts/Build-Plugins.ps1','scripts/pre-commit.ps1')
            $selected = @(Get-FreeTestSelection -Plan $plan -Group Tooling)
            if ($selected[0].Filter -cne 'RequiresExcel=false&RunType!=OnDemand&(Feature=PluginBootstrap|Feature=PluginPublication|Feature=PluginSkillVersion)' -or
                $selected[1].Filter -cne 'RequiresExcel=false&RunType!=OnDemand&(Feature=PreCommit)') {
                throw 'Mixed filters broadened each other.'
            }
            $global:startedOwners = @()
            function Invoke-TestStage {
                param($Project)
                $global:startedOwners += [IO.Path]::GetFileName($Project)
                throw 'selected-root-cause'
            }
            try {
                Invoke-SelectedFreeTests -Plan $plan -Group Tooling -ResultsDirectory fixture
                throw 'Selected failure was swallowed.'
            } catch {
                if ($_ -notmatch 'selected-root-cause') { throw }
            }
            if (($global:startedOwners -join ',') -cne 'ExcelMcp.Packaging.Tests.csproj') { throw 'Execution continued after failure.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task ExplicitLocalOptionsAndContracts_UseNativeFilters()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            . .\scripts\ValidationExecution.ps1
            $plan = Get-ValidationPlan -Paths @()
            $selected = @(Get-FreeTestSelection -Plan $plan -Local -Contracts)
            if (($selected.Owner -join ',') -cne 'Core,CLI,McpServer') { throw 'Contract consumer lost.' }
            foreach ($item in $selected) {
                if ($item.Filter -cne 'RequiresExcel=false&RunType!=OnDemand&(Feature=GeneratedContracts)') {
                    throw 'Contract labels changed.'
                }
            }
            $selected = @(Get-FreeTestSelection -Plan $plan -Local -HookTests -SkillTests -PackagingTests)
            if (($selected.Owner -join ',') -cne 'ScriptSafety,SkillGeneration,Packaging') { throw 'Explicit local owners lost.' }
            if (@(Get-FreeTestSelection -Plan $plan -Local).Count) { throw 'An empty local selection claimed tests.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }
}
