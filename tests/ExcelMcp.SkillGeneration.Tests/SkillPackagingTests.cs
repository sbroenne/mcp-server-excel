using System.IO.Compression;
using Sbroenne.ExcelMcp.Tests.Infrastructure;
using Xunit;
using static Sbroenne.ExcelMcp.Tests.Infrastructure.PackagingScriptTestHelper;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

[Collection("GeneratedAssets")]
[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "SkillGeneration")]
public sealed class SkillPackagingTests
{
    private const string TestVersion = "9.9.9-skillversion";
    private static readonly string BuildAgentSkillsScript = Path.Combine(RepoRoot, "scripts", "Build-AgentSkills.ps1");

    [Fact]
    public async Task BuildAgentSkills_StampsVersionFileIntoEveryPackagedSkill()
    {
        var sandbox = CreateSandbox("agent-skills-version");
        try
        {
            var outputDir = Path.Combine(sandbox, "skills");
            var result = await RunPowerShellFileAsync(
                BuildAgentSkillsScript,
                [
                    "-SkillsDirectory",
                    GeneratedAssetsFixture.SkillsDirectory,
                    "-Version",
                    TestVersion,
                    "-OutputDir",
                    Path.GetRelativePath(RepoRoot, outputDir)
                ]);

            Assert.True(
                result.ExitCode == 0,
                $"Build-AgentSkills.ps1 failed with exit code {result.ExitCode}.{Environment.NewLine}{result.CombinedOutput}");

            var zipPath = Path.Combine(outputDir, $"excel-skills-v{TestVersion}.zip");
            Assert.True(File.Exists(zipPath), $"Agent Skills ZIP was not created: {zipPath}");

            using var archive = ZipFile.OpenRead(zipPath);
            var versionEntries = archive.Entries
                .Where(entry => entry.FullName.EndsWith("/VERSION", StringComparison.Ordinal))
                .OrderBy(entry => entry.FullName, StringComparer.Ordinal)
                .ToList();

            Assert.Equal(
                ["skills/excel-cli-report-formatting/VERSION", "skills/excel-mcp-report-formatting/VERSION"],
                versionEntries.Select(entry => entry.FullName).ToArray());

            foreach (var entry in versionEntries)
            {
                using var reader = new StreamReader(entry.Open());
                Assert.Equal(TestVersion, (await reader.ReadToEndAsync()).Trim());
            }
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

}
