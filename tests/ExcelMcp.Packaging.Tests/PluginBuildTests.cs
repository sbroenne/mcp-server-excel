using Sbroenne.ExcelMcp.Tests.Infrastructure;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Packaging.Tests;

[Collection("GeneratedAssets")]
[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "PluginBootstrap")]
public sealed class PluginBuildTests(ITestOutputHelper output) : PluginTestBase(output)
{
    [Fact]
    public async Task BuildPlugins_ProducesNpxOnlyPackages()
    {
        var sandbox = CreateSandbox("build");
        try
        {
            var outputDirectory = Path.Combine(sandbox, "built-plugins");
            const string version = "9.9.9-test";

            var result = await RunPowerShellFileAsync(
                BuildPluginsScript, ["-Version", version, "-OutputDir", outputDirectory]);

            Assert.True(result.ExitCode == 0, result.CombinedOutput);
            Assert.Contains("[ok] excel-mcp - npx config and skill", result.Stdout, StringComparison.Ordinal);
            Assert.Contains("[ok] excel-cli - argument-safe npx wrapper and skill", result.Stdout, StringComparison.Ordinal);

            var mcpRoot = Path.Combine(outputDirectory, "excel-mcp");
            var cliRoot = Path.Combine(outputDirectory, "excel-cli");
            AssertAgentPluginManifest(mcpRoot, version);
            AssertAgentPluginManifest(cliRoot, version);
            AssertPortableMcpConfiguration(mcpRoot);
            Assert.True(File.Exists(Path.Combine(cliRoot, "bin", "start-cli.ps1")));
            Assert.False(File.Exists(Path.Combine(mcpRoot, "bin", "start-mcp.ps1")));
            Assert.False(File.Exists(Path.Combine(mcpRoot, "bin", "download.ps1")));
            Assert.False(File.Exists(Path.Combine(cliRoot, "bin", "download.ps1")));
            Assert.Empty(Directory.GetFiles(outputDirectory, "install-global.ps1", SearchOption.AllDirectories));

            Assert.True(Directory.Exists(Path.Combine(mcpRoot, "skills", "excel-mcp-report-formatting")),
                "The MCP plugin must contain its matching report-formatting skill.");
            Assert.True(Directory.Exists(Path.Combine(cliRoot, "skills", "excel-cli-report-formatting")),
                "The CLI plugin must contain its matching report-formatting skill.");
            var launcherSkill = Path.Combine(cliRoot, "skills", "excel-cli");
            Assert.True(File.Exists(Path.Combine(launcherSkill, "SKILL.md")),
                "The CLI plugin must expose discovery instructions for ordinary Excel requests.");
            Assert.True(File.Exists(Path.GetFullPath(Path.Combine(launcherSkill, "..", "..", "bin", "start-cli.ps1"))));
            Assert.False(Directory.Exists(Path.Combine(launcherSkill, "references")));
            var skills = Directory.GetDirectories(outputDirectory)
                .Select(plugin => Path.Combine(plugin, "skills"))
                .SelectMany(Directory.GetDirectories).Order(StringComparer.Ordinal).ToArray();
            foreach (var skill in skills)
            {
                AssertSkillDirectoryMatchesSource(
                    Path.Combine(GeneratedAssetsFixture.SkillsDirectory, Path.GetFileName(skill)), skill, version);
            }
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

    [Fact]
    public async Task SyncPublishedPluginRepo_CopiesNpxOnlyPackages()
    {
        var sandbox = CreateSandbox("sync");
        try
        {
            var builtDirectory = Path.Combine(sandbox, "built");
            var publishedDirectory = Directory.CreateDirectory(Path.Combine(sandbox, "published")).FullName;
            File.WriteAllText(Path.Combine(publishedDirectory, "marketplace.json"), "{}");
            const string version = "9.9.10-test";

            var build = await RunPowerShellFileAsync(
                BuildPluginsScript, ["-Version", version, "-OutputDir", builtDirectory]);
            Assert.True(build.ExitCode == 0, build.CombinedOutput);
            var sync = await RunPowerShellFileAsync(
                SyncPublishedRepoScript,
                ["-PublishedRepoDir", publishedDirectory, "-BuiltPluginsDir", builtDirectory, "-Version", version]);

            Assert.True(sync.ExitCode == 0, sync.CombinedOutput);
            Assert.True(File.Exists(Path.Combine(publishedDirectory, "plugins", "excel-mcp", "mcp.json")));
            Assert.True(File.Exists(Path.Combine(publishedDirectory, "plugins", "excel-cli", "bin", "start-cli.ps1")));
            Assert.False(File.Exists(Path.Combine(publishedDirectory, "plugins", "excel-mcp", "bin", "download.ps1")));
            Assert.False(File.Exists(Path.Combine(publishedDirectory, "plugins", "excel-cli", "bin", "download.ps1")));
            Assert.Empty(Directory.GetFiles(publishedDirectory, "install-global.ps1", SearchOption.AllDirectories));

            var validation = await RunPowerShellFileAsync(
                Path.Combine(publishedDirectory, "tests", "Test-Plugins.ps1"), []);
            Assert.True(validation.ExitCode == 0, validation.CombinedOutput);
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

    [Fact]
    public async Task SyncPublishedPluginRepo_RejectsMcpPackageWithoutConfiguration()
    {
        var sandbox = CreateSandbox("sync-missing-mcp-config");
        try
        {
            var builtDirectory = Path.Combine(sandbox, "built");
            var publishedDirectory = Directory.CreateDirectory(Path.Combine(sandbox, "published")).FullName;
            File.WriteAllText(Path.Combine(publishedDirectory, "marketplace.json"), "{}");

            var build = await RunPowerShellFileAsync(
                BuildPluginsScript, ["-Version", "9.9.11-test", "-OutputDir", builtDirectory]);
            Assert.True(build.ExitCode == 0, build.CombinedOutput);
            File.Delete(Path.Combine(builtDirectory, "excel-mcp", "mcp.json"));

            var sync = await RunPowerShellFileAsync(
                SyncPublishedRepoScript,
                ["-PublishedRepoDir", publishedDirectory, "-BuiltPluginsDir", builtDirectory, "-Version", "9.9.11-test"]);

            Assert.NotEqual(0, sync.ExitCode);
            Assert.Contains("excel-mcp is missing mcp.json", sync.Stderr, StringComparison.Ordinal);
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

    [Theory]
    [InlineData("excel-mcp", "bin")]
    [InlineData("excel-mcp", "com.github.copilot")]
    [InlineData("excel-cli", "bin")]
    [InlineData("excel-cli", "com.github.copilot")]
    public async Task SyncPublishedPluginRepo_RejectsRetiredGlobalInstaller(string pluginName, string directory)
    {
        var sandbox = CreateSandbox("sync-retired-installer");
        try
        {
            var builtDirectory = Path.Combine(sandbox, "built");
            var publishedDirectory = Directory.CreateDirectory(Path.Combine(sandbox, "published")).FullName;
            const string version = "9.9.12-test";

            var build = await RunPowerShellFileAsync(
                BuildPluginsScript, ["-Version", version, "-OutputDir", builtDirectory]);
            Assert.True(build.ExitCode == 0, build.CombinedOutput);
            var installerDirectory = Directory.CreateDirectory(Path.Combine(builtDirectory, pluginName, directory)).FullName;
            File.WriteAllText(Path.Combine(installerDirectory, "install-global.ps1"), "# Retired helper");

            var sync = await RunPowerShellFileAsync(
                SyncPublishedRepoScript,
                ["-PublishedRepoDir", publishedDirectory, "-BuiltPluginsDir", builtDirectory, "-Version", version]);

            Assert.NotEqual(0, sync.ExitCode);
            Assert.Contains("Global installation helpers are retired", sync.Stderr, StringComparison.Ordinal);
            Assert.Empty(Directory.GetFileSystemEntries(publishedDirectory));
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }
}
