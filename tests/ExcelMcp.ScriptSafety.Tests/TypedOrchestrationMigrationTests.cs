using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class TypedOrchestrationMigrationTests
{
    [Theory]
    [InlineData("")]
    [InlineData("restore")]
    [InlineData("build")]
    public async Task FocusedBuild_UsesOneMsbuildGraphAndCleansItOnEveryExit(string failure)
    {
        var run = await ValidationSelectionTests.RunAsync($$"""
            . .\scripts\ValidationExecution.ps1
            $global:buildCommands = @()
            $global:buildGraph = $null
            function Invoke-NativeCommand {
                param($Executable, $Arguments)
                if ($Executable -ne 'dotnet') { throw 'Unexpected executable.' }
                $global:buildCommands += ,$Arguments
                $global:buildGraph = $Arguments[1]
                $graph = Get-Content -LiteralPath $global:buildGraph -Raw | ConvertFrom-Json
                if (($graph.solution.projects -join ',') -cne 'src\ExcelMcp.Core\ExcelMcp.Core.csproj,tests\ExcelMcp.Core.Tests\ExcelMcp.Core.Tests.csproj') {
                    throw 'The SDK graph lost selected inputs.'
                }
                if ($Arguments -contains '--no-incremental' -or $Arguments -contains '--disable-build-servers') { throw 'Native optimizations disabled.' }
                if ($Arguments[0] -eq 'build' -and $Arguments -notcontains '-p:UseSharedCompilation=false') {
                    throw 'A persistent compiler server would outlive its owned build.'
                }
                if ($Arguments[0] -eq '{{failure}}') { throw 'native-root-cause' }
                @{ Output='fixture'; Error='' }
            }
            $plan = Get-ValidationPlan -Paths @()
            $plan.BuildProjects = @('src/ExcelMcp.Core/ExcelMcp.Core.csproj','tests/ExcelMcp.Core.Tests/ExcelMcp.Core.Tests.csproj')
            try {
                Invoke-SelectedBuild $plan
                if ('{{failure}}') { throw 'Native failure was hidden.' }
            } catch {
                if (-not '{{failure}}' -or $_ -notmatch 'native-root-cause') { throw }
            }
            if (Test-Path -LiteralPath $global:buildGraph) { throw 'Temporary graph was not removed.' }
            $expected = if ('{{failure}}' -eq 'restore') { 1 } else { 2 }
            if ($global:buildCommands.Count -ne $expected -or $global:buildCommands[0][0] -ne 'restore') { throw 'Wrong build sequence.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task CompleteBuild_HasOneDeterministicOwningGroup()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            . .\scripts\ValidationExecution.ps1
            $plan = Get-ValidationPlan -Full
            foreach ($group in $plan.CiTestGroups) {
                $projects = @(Get-SelectedBuildProjects $plan $group)
                if ($group -eq $plan.FullSolutionBuildGroup) {
                    if (($projects -join ',') -cne 'Sbroenne.ExcelMcp.sln') { throw 'Complete build missing.' }
                } elseif (@($projects | Where-Object { -not $_.EndsWith('.Tests.csproj') }).Count) {
                    throw 'Solution was rebuilt in a second group.'
                }
            }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task ExcelFailure_PreservesPrimaryAndOwnedCleanupErrors()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            . .\scripts\ValidationExecution.ps1
            function Invoke-TestStage { throw 'test-root-cause' }
            $directory = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.Cleanup.$([Guid]::NewGuid().ToString('N'))"
            New-Item -ItemType Directory "$directory\scripts" -Force | Out-Null
            try {
                'param($PipeName); if ($PipeName -notlike "excelmcp-selected-*") { throw "Wrong identity" }; throw "cleanup-root-cause"' |
                    Set-Content "$directory\scripts\Stop-ExcelCliService.ps1"
                $definition = ${function:Invoke-ExcelPlan}.ToString().Replace(
                    '$root = Split-Path -Parent $PSScriptRoot', ('$root = ''' + $directory.Replace("'","''") + ''''))
                Set-Item Function:\Invoke-ExcelPlan ([scriptblock]::Create($definition))
                $plan = Get-ValidationPlan -Paths 'scripts/Test-E2E.ps1'
                $plan.FullE2E = $false
                try {
                    Invoke-ExcelPlan -Plan $plan -ResultsDirectory $directory
                    throw 'Selected failure was swallowed.'
                } catch {
                    if ($_.Exception -isnot [AggregateException] -or $_ -notmatch 'test-root-cause' -or $_ -notmatch 'cleanup-root-cause') { throw }
                }
            } finally { Remove-Item -LiteralPath $directory -Recurse -Force }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task BuiltInventory_UsesWholeFixturesAndOnlySplitsMixedMethodGroups()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            . .\scripts\Get-ExcelTestGroups.ps1
            function Invoke-TestStage {
                param($Environment)
                @(
                    @{Class='Fixture.Cases'; Method='Fixture.Cases.One'; Group='Editing'; OnDemand=$false; Required=$false; Prerequisite=$false},
                    @{Class='Fixture.Cases'; Method='Fixture.Cases.Two'; Group='Editing'; OnDemand=$false; Required=$false; Prerequisite=$false},
                    @{Class='Mixed.Cases'; Method='Mixed.Cases.Normal'; Group='Lifecycle'; OnDemand=$false; Required=$false; Prerequisite=$false},
                    @{Class='Mixed.Cases'; Method='Mixed.Cases.Macro'; Group='VBA'; OnDemand=$false; Required=$false; Prerequisite=$true}
                ) | ConvertTo-Json | Set-Content -LiteralPath $Environment.EXCELMCP_TEST_SELECTION_OUTPUT
            }
            $directory = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.Inventory.$([Guid]::NewGuid().ToString('N'))"
            New-Item -ItemType Directory $directory | Out-Null
            try {
                $groups = @(Get-ExcelTestGroups -ResultsDirectory $directory)
                if ($groups.Count -ne 15) { throw 'An inventory owner or group was lost.' }
                foreach ($owner in @('Core','Service','CLI','McpServer','ComInterop')) {
                    $selected = @($groups | Where-Object ProjectName -eq $owner)
                    if (@($selected | Where-Object { $_.Filter -eq '(FullyQualifiedName~Fixture.Cases.)' -and $_.Cases.Count -eq 2 }).Count -ne 1 -or
                        @($selected | Where-Object { $_.Filter -eq '(FullyQualifiedName=Mixed.Cases.Macro)' -and $_.Cases[0].Prerequisite }).Count -ne 1) {
                        throw 'Fixture boundary or mixed-method group changed.'
                    }
                }
                try { Get-ExcelTestGroups -ResultsDirectory $directory; throw 'Stale inventory accepted.' }
                catch { if ($_ -notmatch 'new inventory path') { throw } }
            } finally { Remove-Item -LiteralPath $directory -Recurse -Force }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }
}
