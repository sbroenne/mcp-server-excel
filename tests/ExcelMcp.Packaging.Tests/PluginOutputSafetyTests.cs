using Xunit;
using static Sbroenne.ExcelMcp.Tests.Infrastructure.PackagingScriptTestHelper;

namespace Sbroenne.ExcelMcp.Packaging.Tests;

[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "PluginSkillVersion")]
public sealed class PluginOutputSafetyTests
{
    private const string TestVersion = "9.9.9-skillversion";
    private static readonly string BuildPluginsScript = Path.Combine(RepoRoot, "scripts", "Build-Plugins.ps1");

    [Fact]
    public async Task BuildPlugins_RequiresExplicitVersion()
    {
        var sandbox = CreateSandbox("plugin-version-required");
        try
        {
            var result = await RunPowerShellFileAsync(
                BuildPluginsScript,
                ["-OutputDir", Path.Combine(sandbox, "built-plugins")]);

            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("Version is required", result.CombinedOutput, StringComparison.Ordinal);
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Fact]
    public async Task BuildPlugins_FailedPreparationPreservesExistingOutput()
    {
        var sandbox = CreateSandbox("plugin-preserves-output");
        var outputDir = Path.Combine(sandbox, "plugins");
        try
        {
            Directory.CreateDirectory(outputDir);
            var prior = Path.Combine(outputDir, "excel-cli");
            Directory.CreateDirectory(prior);
            File.WriteAllText(Path.Combine(prior, "prior.txt"), "last good");
            File.WriteAllText(Path.Combine(outputDir, "unrelated.txt"), "keep");
            var result = await RunPowerShellFileAsync(BuildPluginsScript,
                ["-Version", TestVersion, "-OutputDir", outputDir, "-SkillsDirectory", Path.Combine(sandbox, "missing")]);
            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("missing", result.CombinedOutput, StringComparison.Ordinal);
            Assert.Equal("last good", File.ReadAllText(Path.Combine(prior, "prior.txt")));
            Assert.Equal("keep", File.ReadAllText(Path.Combine(outputDir, "unrelated.txt")));
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

}
