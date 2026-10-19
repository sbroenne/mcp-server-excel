using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class ServiceTestClassSelectionRegressionTests
{
    [Fact]
    public async Task PersistentServicePartialFiles_SelectOnlyTheirCompleteTestClass()
    {
        var run = await ValidationSelectionTests.RunAsync("""
            $paths = @(
                'tests/ExcelMcp.Service.Tests/PersistentServiceWindowTests.Arrange.cs',
                'tests/ExcelMcp.Service.Tests/PersistentServiceWindowTests.Visibility.cs')
            $plan = Get-ValidationPlan -Paths $paths
            if ($plan.ExcelSelections.Count -ne 1) { throw "Expected one deduplicated class selection, got $($plan.ExcelSelections.Count)." }
            $selection = $plan.ExcelSelections[0]
            if ($selection.Project -cne 'Service' -or
                $selection.Filter -cne 'FullyQualifiedName~Sbroenne.ExcelMcp.Service.Tests.PersistentServiceWindowTests' -or
                $selection.Area -cne 'PersistentServiceWindowTests') {
                throw "Wrong Service test selection: $($selection | ConvertTo-Json -Compress)"
            }
            if ($plan.FastProjects.Count -or $plan.ProcessProjects.Count -or $plan.ToolingProjects.Count) {
                throw 'A Service class change selected unrelated test projects.'
            }
            if (-not $plan.BuildOnly) { throw 'The selected Service build was not assigned to the build-only CI job.' }
            if (($plan.ExcelGroups -join ',') -cne 'PersistentServiceWindowTests') {
                throw "Wrong Excel groups: $($plan.ExcelGroups -join ',')"
            }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Theory]
    [InlineData("PersistentServiceCleanupFailureTests.cs")]
    [InlineData("PersistentServicePowerQueryRefreshPivotTests.cs")]
    [InlineData("PersistentServiceScreenshotCaptureTests.Advanced.cs")]
    [InlineData("PersistentServiceOwnedContextVisibilityTests.cs")]
    [InlineData("PersistentServiceVbaFailureTests.cs")]
    public async Task OtherServiceTestFiles_KeepOwningProjectSelection(string file)
    {
        var run = await ValidationSelectionTests.RunAsync($$"""
            $plan = Get-ValidationPlan -Paths 'tests/ExcelMcp.Service.Tests/{{file}}'
            if (($plan.FastProjects -join ',') -cne 'Service' -or
                $plan.ExcelSelections.Count -ne 1 -or
                $plan.ExcelSelections[0].Project -cne 'Service' -or
                $plan.ExcelSelections[0].Filter -cne 'FullyQualifiedName~Sbroenne.ExcelMcp.' -or
                $plan.ExcelSelections[0].Area -cne 'All') {
                throw "Existing owning-project coverage changed for {{file}}."
            }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }
}
