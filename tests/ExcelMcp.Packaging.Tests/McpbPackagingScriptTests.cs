using System.IO.Compression;
using System.Text.Json;
using Sbroenne.ExcelMcp.Build;
using Xunit;
using static Sbroenne.ExcelMcp.Tests.Infrastructure.PackagingScriptTestHelper;

namespace Sbroenne.ExcelMcp.Packaging.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "McpbPackaging")]
public sealed class McpbPackagingScriptTests
{
    private static readonly string[] BundleEntries = ["CHANGELOG.md", "LICENSE", "README.md", "icon-512.png", "manifest.json"];
    private static readonly string[] ServerArguments = ["-y", "@sbroenne/mcp-server-excel@latest"];
    [Fact]
    public async Task Build_CreatesMetadataOnlyBundleWithDirectNpxLatest()
    {
        var sandbox = CreateSandbox("mcpb");
        try
        {
            var bundle = StageMcpbInputs(sandbox);
            var result = await RunPowerShellFileAsync(Path.Combine(bundle, "Build-McpBundle.ps1"), ["-Version", "1.2.3"]);
            Assert.True(result.ExitCode == 0, result.CombinedOutput);
            await AssertDirectNpxBundleAsync(Path.Combine(bundle, "artifacts", "excel-mcp-1.2.3.mcpb"));
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

    [Fact]
    public async Task AggregatePackaging_McpbCreatesMetadataOnlyBundleWithoutPublishingRuntime()
    {
        var sandbox = CreateSandbox("mcpb-aggregate");
        try
        {
            StageMcpbInputs(sandbox);
            var output = Path.Combine(sandbox, "artifacts", "packages");
            var commands = new PackageCommands(_ => throw new InvalidOperationException("Metadata bundles must not install npm dependencies or publish a runtime."));
            await new PackageExecution(sandbox, commands).ReleaseAsync(new PackageOptions
            {
                Components = ["Mcpb"],
                Version = "1.2.3",
                OutputDirectory = output
            });
            await AssertDirectNpxBundleAsync(Path.Combine(output, "mcpb", "excel-mcp-1.2.3.mcpb"));
            Assert.False(Directory.Exists(Path.Combine(output, "runtimes")));
            Assert.False(Directory.Exists(Path.Combine(output, "nuget")));
            Assert.False(Directory.Exists(Path.Combine(output, "npm")));
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

    [Theory]
    [InlineData("wrong-version")]
    [InlineData("")]
    public async Task AggregatePackaging_RejectsMismatchedPreparedSkillsBeforeCreatingOutput(string stamp)
    {
        var sandbox = CreateSandbox("prepared-skill-version");
        try
        {
            var skills = Path.Combine(sandbox, "prepared");
            foreach (var name in new[] { "excel-cli", "excel-mcp" })
            {
                var directory = Directory.CreateDirectory(Path.Combine(skills, name)).FullName;
                if (stamp.Length > 0) { File.WriteAllText(Path.Combine(directory, "VERSION"), stamp); }
            }
            var output = Path.Combine(sandbox, "artifacts", "packages");
            var commands = new PackageCommands(_ => throw new InvalidOperationException("Payload packaging was reached."));
            var error = await Assert.ThrowsAsync<InvalidOperationException>(() => new PackageExecution(sandbox, commands).ReleaseAsync(new PackageOptions
            {
                Components = ["Plugins"],
                Version = "1.2.3",
                SkillsDirectory = skills,
                OutputDirectory = output
            }));
            Assert.Contains("skill must match package version", error.Message, StringComparison.Ordinal);
            Assert.False(Directory.Exists(output));
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

    [Fact]
    public void FailedBuild_PreservesExistingOutput()
    {
        var sandbox = CreateSandbox("mcpb-preserved-output");
        try
        {
            var bundle = StageMcpbInputs(sandbox);
            var output = Directory.CreateDirectory(Path.Combine(bundle, "artifacts")).FullName;
            var previous = Path.Combine(output, "excel-mcp-1.2.3.mcpb");
            var unrelated = Path.Combine(output, "keep.txt");
            File.WriteAllText(previous, "previous-good-package");
            File.WriteAllText(unrelated, "unrelated");
            var error = Assert.Throws<IOException>(() => new AuthoredPackages(sandbox, (_, _) =>
                throw new IOException("archive-root-cause")).Mcpb("1.2.3", output));
            Assert.Contains("archive-root-cause", error.Message, StringComparison.Ordinal);
            Assert.Equal("previous-good-package", File.ReadAllText(previous));
            Assert.Equal("unrelated", File.ReadAllText(unrelated));
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

    [Theory]
    [InlineData("src/output")]
    [InlineData("skills/generated")]
    [InlineData(".github/plugins/output")]
    public void PackageOutputs_RejectSourceDestinations(string relativePath)
    {
        var sandbox = CreateSandbox("package-source-destination");
        try
        {
            var output = Path.Combine(sandbox, relativePath);
            var error = Assert.Throws<ArgumentException>(() => PackageFiles.AssertOutput(output, sandbox));
            Assert.Contains("Unsafe package output", error.Message, StringComparison.Ordinal);
            Assert.False(Directory.Exists(output));
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

    [Fact]
    public async Task PublicationSync_RejectsItsOwnSourceTreeBeforeInspectingPayload()
    {
        var sandbox = CreateSandbox("publication-source-tree");
        try
        {
            var scripts = StagePublicationScripts(sandbox);
            var built = Directory.CreateDirectory(Path.Combine(sandbox, "built")).FullName;
            var result = await RunPowerShellFileAsync(Path.Combine(scripts, "Sync-PublishedPluginRepo.ps1"),
                ["-PublishedRepoDir", sandbox, "-BuiltPluginsDir", built, "-Version", "1.2.3"]);
            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("Unsafe package output", result.Stderr, StringComparison.Ordinal);
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

    [Theory]
    [InlineData("version.txt")]
    [InlineData("skills/excel-cli-report-formatting/SKILL.md")]
    public async Task PublicationSync_IncompletePayloadCannotReplaceExistingOutput(string missingFile)
    {
        var sandbox = CreateSandbox("publication-incomplete-payload");
        try
        {
            var scripts = StagePublicationScripts(sandbox);
            var overlay = Directory.CreateDirectory(Path.Combine(sandbox, ".github", "plugins", "marketplace-repo")).FullName;
            File.WriteAllText(Path.Combine(overlay, "README.md"), "new overlay");
            var built = Path.Combine(sandbox, "artifacts", "built");
            foreach (var name in new[] { "excel-cli", "excel-mcp" })
            {
                var plugin = Directory.CreateDirectory(Path.Combine(built, name)).FullName;
                var manifest = File.ReadAllText(Path.Combine(RepoRoot, ".github", "plugins", name, "plugin.json"));
                File.WriteAllText(Path.Combine(plugin, "plugin.json"), manifest.Replace("0.0.0", "1.2.3", StringComparison.Ordinal));
                foreach (var file in new[] {
                    "README.md", "version.txt", $"skills/{name}-report-formatting/SKILL.md",
                    $"skills/{name}-report-formatting/VERSION", $"skills/{name}-report-formatting/references/report-formatting.md"
                })
                {
                    var destination = Path.Combine(plugin, file);
                    Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
                    File.WriteAllText(destination, "1.2.3");
                }
                if (name == "excel-cli")
                {
                    var wrapper = Path.Combine(plugin, "bin", "start-cli.ps1");
                    Directory.CreateDirectory(Path.GetDirectoryName(wrapper)!);
                    File.WriteAllText(wrapper, "1.2.3");
                }
                else { File.Copy(Path.Combine(RepoRoot, ".github", "plugins", "excel-mcp", "mcp.json"), Path.Combine(plugin, "mcp.json")); }
            }
            File.Delete(Path.Combine(built, "excel-cli", missingFile));
            var output = Directory.CreateDirectory(Path.Combine(sandbox, "artifacts", "published")).FullName;
            File.WriteAllText(Path.Combine(output, "README.md"), "previous output");
            var result = await RunPowerShellFileAsync(Path.Combine(scripts, "Sync-PublishedPluginRepo.ps1"),
                ["-PublishedRepoDir", output, "-BuiltPluginsDir", built, "-Version", "1.2.3"]);
            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("Incomplete plugin payload", result.Stderr, StringComparison.Ordinal);
            Assert.Equal("previous output", File.ReadAllText(Path.Combine(output, "README.md")));
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

    private static string StagePublicationScripts(string sandbox)
    {
        var scripts = Directory.CreateDirectory(Path.Combine(sandbox, "scripts")).FullName;
        foreach (var name in new[] { "Sync-PublishedPluginRepo.ps1", "PackageHelpers.ps1" })
        {
            File.Copy(Path.Combine(RepoRoot, "scripts", name), Path.Combine(scripts, name));
        }
        return scripts;
    }

    private static string StageMcpbInputs(string sandbox)
    {
        var bundle = Directory.CreateDirectory(Path.Combine(sandbox, "mcpb")).FullName;
        var scripts = Directory.CreateDirectory(Path.Combine(sandbox, "scripts")).FullName;
        foreach (var name in new[] { "Build-ReleasePackages.ps1", "PackageHelpers.ps1" })
        {
            File.Copy(Path.Combine(RepoRoot, "scripts", name), Path.Combine(scripts, name));
        }
        foreach (var name in new[] { "Build-McpBundle.ps1", "manifest.json", "README.md", "icon-512.png" })
        {
            File.Copy(Path.Combine(RepoRoot, "mcpb", name), Path.Combine(bundle, name));
        }
        foreach (var name in new[] { "LICENSE", "CHANGELOG.md" }) { File.Copy(Path.Combine(RepoRoot, name), Path.Combine(sandbox, name)); }
        return bundle;
    }

    private static async Task AssertDirectNpxBundleAsync(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        Assert.Equal(BundleEntries,
            archive.Entries.Select(entry => entry.FullName).Order(StringComparer.Ordinal));
        var entry = archive.GetEntry("manifest.json");
        Assert.NotNull(entry);
        using var stream = entry.Open();
        using var manifest = await JsonDocument.ParseAsync(stream);
        Assert.Equal("1.2.3", manifest.RootElement.GetProperty("version").GetString());
        var server = manifest.RootElement.GetProperty("server");
        Assert.Equal("node", server.GetProperty("type").GetString());
        Assert.Equal("@sbroenne/mcp-server-excel", server.GetProperty("entry_point").GetString());
        var config = server.GetProperty("mcp_config");
        Assert.Equal("npx", config.GetProperty("command").GetString());
        Assert.Equal(ServerArguments,
            config.GetProperty("args").EnumerateArray().Select(argument => argument.GetString()));
        Assert.Empty(config.GetProperty("env").EnumerateObject());
    }
}
