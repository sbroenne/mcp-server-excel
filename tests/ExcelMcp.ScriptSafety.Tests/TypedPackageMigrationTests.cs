using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class TypedPackageMigrationTests
{
    [Theory]
    [InlineData("PackageHelpers.ps1")]
    [InlineData("ExtensionPackaging.ps1")]
    public async Task SharedPackageLogic_SelectsPackageAndSkillConsumersWithoutExcel(string file)
    {
        var run = await ValidationSelectionTests.RunAsync($$"""
            $plan = Get-ValidationPlan -Paths 'scripts/{{file}}'
            if ('Packaging' -notin $plan.ToolingProjects -or 'SkillGeneration' -notin $plan.ToolingProjects -or
                $plan.Excel -or $plan.FullE2E) { throw 'Wrong package helper consumers.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

}
