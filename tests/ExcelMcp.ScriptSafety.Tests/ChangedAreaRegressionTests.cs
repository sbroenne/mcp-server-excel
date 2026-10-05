using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class ChangedAreaRegressionTests
{
    [Theory]
    [InlineData("Range", "RangeCommands.Values.cs")]
    [InlineData("PowerQuery", "PowerQueryCommands.Read.cs")]
    public async Task ImplementationChange_SelectsItsFeatureWithoutUnrelatedAdapters(string area, string file)
    {
        var run = await ValidationSelectionTests.RunAsync($$"""
            $plan = Get-ValidationPlan -Paths 'src/ExcelMcp.Core/Commands/{{area}}/{{file}}'
            if (($plan.FastProjects -join ',') -ne 'Core') { throw 'Unrelated runtime projects selected.' }
            if ($plan.ProcessProjects.Count) { throw 'Unchanged CLI process behavior selected.' }
            if ($plan.NpmTests) { throw 'Unchanged npm launcher tests selected.' }
            if (-not $plan.FastFilters.Core.Contains('{{area}}')) { throw 'Owning Core feature filter missing.' }
            if (-not @($plan.ExcelSelections | Where-Object { $_.Project -eq 'Service' -and $_.Area -eq '{{area}}' }).Count) {
                throw 'Owning Service workbook cases missing.'
            }
            if ($plan.FullE2E) { throw 'Unchanged full acceptance workflow selected.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task TestOnlyChange_SelectsItsClassRatherThanItsWholeProject()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            $plan = Get-ValidationPlan -Paths 'tests/ExcelMcp.Core.Tests/Unit/GeneratedActionContractTests.cs'
            if (($plan.FastProjects -join ',') -ne 'Core') { throw 'Unrelated project selected.' }
            if ($plan.FastFilters.Core -ne 'FullyQualifiedName~Sbroenne.ExcelMcp.Core.Tests.Unit.GeneratedActionContractTests.') {
                throw 'Test-only selection was broadened beyond its class.'
            }
            if ($plan.Excel -or $plan.ProcessProjects.Count) { throw 'Unrelated integration checks selected.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task UnknownInput_FailsInsteadOfSelectingEverything()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            Get-ValidationPlan -Paths 'unknown-build-input.config'
            """);
        Assert.NotEqual(0, run.ExitCode);
        Assert.Contains("No validation mapping", run.Output, StringComparison.Ordinal);
    }
}
