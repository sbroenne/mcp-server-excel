using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class TypedPackageMigrationTests
{
    [Theory]
    [InlineData("Build-AgentSkills.ps1")]
    [InlineData("Build-Plugins.ps1")]
    [InlineData("Build-NpmPackages.ps1")]
    [InlineData("Test-NpmPackages.ps1")]
    [InlineData("Build-ReleasePackages.ps1")]
    public void PackageScripts_DoNotBootstrapTheDevelopmentTool(string name)
    {
        var source = File.ReadAllText(Path.Combine(TypedValidationPolicyTests.Root, "scripts", name));
        Assert.Contains("PackageHelpers.ps1", source, StringComparison.Ordinal);
        Assert.DoesNotContain("Invoke-TypedPackage", source, StringComparison.Ordinal);
        Assert.DoesNotContain("build.ps1", source, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("EXCELMCP_BUILD_ROOT", source, StringComparison.Ordinal);
    }

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

    [Fact]
    public async Task RemovedAlias_DoesNotSelectUnchangedWorkbookAcceptance()
    {
        Assert.False(File.Exists(Path.Combine(TypedValidationPolicyTests.Root, "scripts", "Test-CliApiCoverage.ps1")));
        var run = await ValidationSelectionTests.RunAsync("""
            $plan = Get-ValidationPlan -Paths 'scripts/Test-CliApiCoverage.ps1'
            if ('ScriptSafety' -notin $plan.ToolingProjects -or $plan.Excel -or $plan.FullE2E) { throw 'Removed alias selected acceptance.' }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }
}
