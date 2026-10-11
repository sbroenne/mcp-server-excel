using Sbroenne.ExcelMcp.Build;
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
    public void PackageCommands_ForwardWithoutDuplicatingImplementation(string name)
    {
        var source = File.ReadAllText(Path.Combine(TypedValidationPolicyTests.Root, "scripts", name));
        Assert.Contains("Invoke-TypedPackage", source, StringComparison.Ordinal);
        Assert.DoesNotContain("Compress-Archive", source, StringComparison.Ordinal);
        Assert.DoesNotContain("dotnet publish", source, StringComparison.Ordinal);
        Assert.DoesNotContain("npm.cmd", source, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("PackageFiles.cs")]
    [InlineData("AuthoredPackages.cs")]
    [InlineData("RuntimePackages.cs")]
    [InlineData("PackageExecution.cs")]
    public void PackageLogic_SelectsPackageAndSkillConsumersWithoutExcel(string file)
    {
        var plan = new ValidationPolicy(TypedValidationPolicyTests.Root).Select([$"tools/ExcelMcp.Build/{file}"]);
        Assert.Contains("Packaging", plan.ToolingProjects);
        Assert.Contains("SkillGeneration", plan.ToolingProjects);
        Assert.False(plan.Excel);
        Assert.False(plan.FullE2E);
        Assert.True(plan.Cli && plan.Mcp && plan.Extension && plan.Mcpb && plan.Skills && plan.Plugins);
    }
}
