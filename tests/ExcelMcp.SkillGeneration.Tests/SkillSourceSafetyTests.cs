using Xunit;
using static Sbroenne.ExcelMcp.Tests.Infrastructure.PackagingScriptTestHelper;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "SkillGeneration")]
public sealed class SkillSourceSafetyTests
{
    private static readonly string BuildAgentSkillsScript = Path.Combine(RepoRoot, "scripts", "Build-AgentSkills.ps1");

    [Fact]
    public async Task GenerateSkills_RejectsSourceOutputWithoutChangingSources()
    {
        var output = Path.Combine(RepoRoot, "skills");
        var result = await RunPowerShellFileAsync(BuildAgentSkillsScript,
            ["-GenerateOnly", "-OutputDir", output]);
        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("Unsafe package output directory", result.CombinedOutput, StringComparison.Ordinal);
        Assert.True(File.Exists(Path.Combine(output, "excel-mcp-report-formatting", "SKILL.md")));
    }

    [Fact]
    public void CanonicalSkillSources_DoNotContainGeneratedVersionFiles()
    {
        var versionFiles = Directory
            .GetDirectories(Path.Combine(RepoRoot, "skills"), "excel-*")
            .Select(skillDirectory => Path.Combine(skillDirectory, "VERSION"));

        Assert.All(
            versionFiles,
            versionFile => Assert.False(
                File.Exists(versionFile),
                $"Canonical skill source contains generated package metadata: {versionFile}"));
    }

    [Fact]
    public async Task BuildAgentSkills_RequiresExplicitVersion()
    {
        var sandbox = CreateSandbox("agent-skills-version-required");
        try
        {
            var result = await RunPowerShellFileAsync(
                BuildAgentSkillsScript,
                ["-OutputDir", Path.GetRelativePath(RepoRoot, Path.Combine(sandbox, "skills"))]);

            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("Version is required", result.CombinedOutput, StringComparison.Ordinal);
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }
}
