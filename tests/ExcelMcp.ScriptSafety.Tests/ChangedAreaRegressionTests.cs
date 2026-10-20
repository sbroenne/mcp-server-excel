using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class ChangedAreaRegressionTests
{
    [Fact]
    public async Task TestChanges_SelectTheOwningGroupIncludingDeletedFiles()
    {
        var run = await ValidationSelectionTests.RunCasesAsync(
        [
            ["tests/ExcelMcp.Core.Tests/Unit/GeneratedActionContractTests.cs", "Core", "false"],
            ["tests/ExcelMcp.Core.Tests/Unit/DeletedTests.More.cs", "Core", "false"],
            ["tests/ExcelMcp.McpServer.Tests/Integration/Tools/CalculationGuidanceContractTests.cs", "McpServer", "false"],
            ["tests/ExcelMcp.McpServer.Tests/Integration/Tools/McpProcessOwnershipTests.cs", "McpServer", "true"],
            ["tests/ExcelMcp.McpServer.Tests/Integration/Tools/McpSessionLifetimeTests.cs", "McpServer", "true"],
            ["tests/ExcelMcp.McpServer.Tests/Integration/Tools/McpOwnedExcelProcesses.cs", "McpServer", "true"],
            ["tests/ExcelMcp.McpServer.Tests/Integration/Tools/McpIntegrationTestBase.cs", "McpServer", "true"],
            ["tests/ExcelMcp.McpServer.Tests/Integration/Tools/McpProgramTransportFixture.cs", "McpServer", "true"],
            ["tests/ExcelMcp.McpServer.Tests/Integration/Tools/McpResponseAssertions.cs", "McpServer", "true"],
            ["tests/ExcelMcp.McpServer.Tests/Integration/Tools/ProgramTransportTestHost.cs", "McpServer", "true"],
            ["tests/ExcelMcp.McpServer.Tests/Integration/Tools/McpProcessOwnershipTests.Deleted.cs", "McpServer", "true"],
            ["tests/ExcelMcp.Service.Tests/DeletedFixture.cs", "Service", "true"]
        ], """
            $path, $owner, $excel = $case
            $plan = Get-ValidationPlan -Paths $path
            if (($plan.FastProjects -join ',') -cne $owner -or
                $plan.Excel -ne [bool]::Parse($excel) -or $plan.FullE2E) { throw 'Wrong owning test group.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task RealExcelTransportFiles_CannotBeClassifiedAsRecordingOnlyProtocols()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            $checked = 0
            foreach ($file in Get-ChildItem .\tests\ExcelMcp.McpServer.Tests\Integration\Tools -Filter *.cs) {
                if ((Get-Content $file.FullName -Raw) -notmatch '\[Trait\("RequiresExcel",\s*"true"\)\]') { continue }
                $checked++
                $plan = Get-ValidationPlan -Paths "tests/ExcelMcp.McpServer.Tests/Integration/Tools/$($file.Name)"
                if (-not $plan.Excel -or $plan.ExcelSelections.Count -ne 1 -or
                    $plan.ExcelSelections[0].Project -ne 'McpServer') {
                    throw "Real-Excel transport coverage missing for $($file.Name)."
                }
            }
            if ($checked -lt 5) { throw 'The real-Excel transport fixture check did not inspect its expected inputs.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task Contracts_SelectBothGeneratedAdaptersAndAreaLabels()
    {
        var run = await ValidationSelectionTests.RunCasesAsync(
        [
            ["Range", "IRangeCommands.cs"],
            ["Chart", "IChartConfigCommands.cs"]
        ], """
            $area, $contract = $case
            $plan = Get-ValidationPlan -Paths "src/ExcelMcp.Core/Commands/$area/$contract"
            foreach ($owner in @('Core','CLI','McpServer')) {
                if ($plan.FastFilters[$owner] -notmatch 'Feature=GeneratedContracts') { throw "Generated consumer lost: $owner" }
            }
            foreach ($owner in @('CLI','McpServer')) {
                if ($plan.FastFilters[$owner] -notmatch "Feature=$area") { throw "Argument mapping lost: $owner" }
            }
            if (-not $plan.FullSolutionBuild -or -not $plan.DocumentationCounts -or $plan.FullE2E) {
                throw 'Contract build ownership changed.'
            }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task SharedSource_UsesExplicitRuntimeCoverageInsteadOfRebuildingASourceCatalogue()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            $plan = Get-ValidationPlan -Paths 'src/ExcelMcp.Core/DataModel/DeletedHelper.cs'
            if (-not $plan.FullE2E -or -not $plan.InfrastructureDiagnostics -or -not $plan.FullSolutionBuild) {
                throw 'Shared runtime consumers missing.'
            }
            if (($plan.FastProjects -join ',') -cne 'CLI,ComInterop,Core,McpServer,Service') { throw 'Runtime owner lost.' }
            if ($plan.ToolingProjects.Count) { throw 'Runtime change selected publication tests.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Theory]
    [InlineData("invalid-target", "Unknown build command")]
    [InlineData("plan --unknown", "Unknown option")]
    [InlineData("plan --full --full", "Duplicate option")]
    [InlineData("plan --base-ref", "missing value")]
    [InlineData("plan --area ../Range", "Area must")]
    public async Task InvalidCommands_PreserveTheirDiagnostic(string arguments, string diagnostic)
    {
        var run = await ValidationSelectionTests.RunAsync($"& .\\build.ps1 {arguments}");
        Assert.NotEqual(0, run.ExitCode);
        Assert.Contains(diagnostic, run.Output, StringComparison.Ordinal);
    }
}
