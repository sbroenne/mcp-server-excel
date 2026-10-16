using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class ChangedAreaRegressionTests
{
    [Theory]
    [InlineData("tests/ExcelMcp.Core.Tests/Unit/GeneratedActionContractTests.cs", "Core", false)]
    [InlineData("tests/ExcelMcp.Core.Tests/Unit/DeletedTests.More.cs", "Core", false)]
    [InlineData("tests/ExcelMcp.McpServer.Tests/Integration/Tools/CalculationGuidanceContractTests.cs", "McpServer", false)]
    [InlineData("tests/ExcelMcp.Service.Tests/DeletedFixture.cs", "Service", true)]
    public async Task TestChanges_SelectTheOwningGroupIncludingDeletedFiles(string path, string owner, bool excel)
    {
        var run = await ValidationSelectionTests.RunAsync($$"""
            $plan = Get-ValidationPlan -Paths '{{path}}'
            if (($plan.FastProjects -join ',') -cne '{{owner}}' -or
                $plan.Excel -ne {{(excel ? "$true" : "$false")}} -or $plan.FullE2E) { throw 'Wrong owning test group.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Theory]
    [InlineData("Range", "IRangeCommands.cs")]
    [InlineData("Chart", "IChartConfigCommands.cs")]
    public async Task Contracts_SelectBothGeneratedAdaptersAndAreaLabels(string area, string contract)
    {
        var run = await ValidationSelectionTests.RunAsync($$"""
            $plan = Get-ValidationPlan -Paths 'src/ExcelMcp.Core/Commands/{{area}}/{{contract}}'
            foreach ($owner in @('Core','CLI','McpServer')) {
                if ($plan.FastFilters[$owner] -notmatch 'Feature=GeneratedContracts') { throw "Generated consumer lost: $owner" }
            }
            foreach ($owner in @('CLI','McpServer')) {
                if ($plan.FastFilters[$owner] -notmatch 'Feature={{area}}') { throw "Argument mapping lost: $owner" }
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
