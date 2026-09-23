using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

/// <summary>
/// Integration tests for Agent Plugins 1.0 packaging and runtime bootstrap flows.
/// These exercise the real PowerShell build/sync scripts against canonical source templates
/// and isolated output repositories without touching real user state.
/// </summary>
public sealed class PluginBootstrapBuildTests
{
    private const string AgentPluginSchema = "https://agent-plugins.org/schemas/1.0.0/plugin.schema.json";
    private const string AgentPluginMcpSchema = "https://agent-plugins.org/schemas/1.0.0/mcp.schema.json";
    private static readonly string RepoRoot = FindRepoRoot();
    private static readonly string BuildAgentSkillsScript = Path.Combine(RepoRoot, "scripts", "Build-AgentSkills.ps1");
    private static readonly string BuildPluginsScript = Path.Combine(RepoRoot, "scripts", "Build-Plugins.ps1");
    private static readonly string SyncPublishedRepoScript = Path.Combine(RepoRoot, "scripts", "Sync-PublishedPluginRepo.ps1");

    [Theory]
    [InlineData("excel-mcp")]
    [InlineData("excel-cli")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "AgentPluginSpec")]
    public void SourcePluginManifest_ConformsToAgentPluginsV1(string pluginName)
    {
        var pluginRoot = Path.Combine(RepoRoot, ".github", "plugins", pluginName);

        AssertAgentPluginManifest(pluginRoot, expectedVersion: "0.0.0");
        AssertAgentSkill(Path.Combine(RepoRoot, "skills", pluginName), pluginName);
        Assert.True(File.Exists(Path.Combine(pluginRoot, "com.github.copilot", "bin", "install-global.ps1")));
        Assert.False(File.Exists(Path.Combine(pluginRoot, "bin", "install-global.ps1")));
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "AgentPluginSpec")]
    public void ExcelMcpSource_UsesPortableMcpConfiguration()
    {
        var pluginRoot = Path.Combine(RepoRoot, ".github", "plugins", "excel-mcp");

        AssertPortableMcpConfiguration(pluginRoot);
        Assert.False(File.Exists(Path.Combine(pluginRoot, ".mcp.json")));
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "AgentPluginSpec")]
    public void PublishWorkflow_ValidatesBothSkillsWithPinnedOfficialSkillsRef()
    {
        var workflowPath = Path.Combine(RepoRoot, ".github", "workflows", "publish-plugins.yml");
        var content = File.ReadAllText(workflowPath);

        Assert.Contains("actions/setup-python@v6", content);
        Assert.Contains(
            "git+https://github.com/agentskills/agentskills.git@69ef37e9424c0a7ea9dd2293b559e43ec8176379#subdirectory=skills-ref",
            content);
        Assert.Contains(@"skills-ref validate source\plugins\excel-mcp\skills\excel-mcp", content);
        Assert.Contains(@"skills-ref validate source\plugins\excel-cli\skills\excel-cli", content);
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public void BuildAgentSkills_UsesPortableNewlinesForCliSyntaxNotice()
    {
        var script = File.ReadAllText(BuildAgentSkillsScript);

        Assert.Contains("-replace \"`r`n?\", \"`n\"", script, StringComparison.Ordinal);
        Assert.Contains("$cliSyntaxNotice`n`n$sourceContent", script, StringComparison.Ordinal);
        Assert.DoesNotContain("$cliSyntaxNotice`r`n", script, StringComparison.Ordinal);
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public void ReleaseWorkflow_PublishesSha256SumsForWindowsRuntimeArchives()
    {
        var workflowPath = Path.Combine(RepoRoot, ".github", "workflows", "release.yml");
        var content = File.ReadAllText(workflowPath);

        Assert.Contains("SHA256SUMS", content, StringComparison.Ordinal);
        Assert.Contains("Get-FileHash", content, StringComparison.Ordinal);
        Assert.Contains("artifacts/mcp-server-zip/*.zip", content, StringComparison.Ordinal);
        Assert.Contains("artifacts/cli-zip/*.zip", content, StringComparison.Ordinal);
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task BuildPlugins_PreservesCurrentBootstrapAssetsAndDropsLegacyBootstrapFiles()
    {
        var sandbox = CreateSandbox("build-preserves-bootstrap-assets");
        try
        {
            var outputDir = Path.Combine(sandbox, "built-plugins");
            var version = "9.9.9-test";

            var result = await RunPowerShellFileAsync(
                BuildPluginsScript,
                [
                    "-Version", version,
                    "-OutputDir", outputDir
                ]);

            Assert.Equal(0, result.ExitCode);

            AssertBootstrapAssetSet(
                Path.Combine(outputDir, "excel-mcp"),
                "mcp.json",
                "bin/start-mcp.ps1",
                "bin/download.ps1",
                "com.github.copilot/bin/install-global.ps1");

            AssertBootstrapAssetSet(
                Path.Combine(outputDir, "excel-cli"),
                "bin/start-cli.ps1",
                "bin/download.ps1",
                "com.github.copilot/bin/install-global.ps1");

            AssertBootstrapAssetsAbsent(
                Path.Combine(outputDir, "excel-mcp"),
                ".mcp.json",
                "bin/download-mcp.ps1",
                "bin/bootstrap-state.json");

            AssertBootstrapAssetsAbsent(
                Path.Combine(outputDir, "excel-cli"),
                "bin/download-cli.ps1",
                "bin/bootstrap-state.json");

            Assert.False(File.Exists(Path.Combine(outputDir, "excel-mcp", "bin", "mcp-excel.exe")));
            Assert.False(File.Exists(Path.Combine(outputDir, "excel-cli", "bin", "excelcli.exe")));
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task BuildPlugins_RefreshesVersionAndSkillContentWithoutClobberingCliOverlay()
    {
        var sandbox = CreateSandbox("build-refreshes-version-and-skills");
        try
        {
            var outputDir = Path.Combine(sandbox, "built-plugins");
            var version = "9.9.10-test";

            var result = await RunPowerShellFileAsync(
                BuildPluginsScript,
                [
                    "-Version", version,
                    "-OutputDir", outputDir
                ]);

            Assert.Equal(0, result.ExitCode);

            Assert.Equal(
                version,
                File.ReadAllText(Path.Combine(outputDir, "excel-mcp", "version.txt")).Trim());
            Assert.Equal(
                version,
                File.ReadAllText(Path.Combine(outputDir, "excel-cli", "version.txt")).Trim());

            using var mcpPluginJson = JsonDocument.Parse(File.ReadAllText(Path.Combine(outputDir, "excel-mcp", "plugin.json")));
            using var cliPluginJson = JsonDocument.Parse(File.ReadAllText(Path.Combine(outputDir, "excel-cli", "plugin.json")));
            Assert.Equal(version, mcpPluginJson.RootElement.GetProperty("version").GetString());
            Assert.Equal(version, cliPluginJson.RootElement.GetProperty("version").GetString());
            AssertAgentPluginManifest(Path.Combine(outputDir, "excel-mcp"), version);
            AssertAgentPluginManifest(Path.Combine(outputDir, "excel-cli"), version);
            AssertPortableMcpConfiguration(Path.Combine(outputDir, "excel-mcp"));

            var sourceMcpSkill = File.ReadAllText(Path.Combine(RepoRoot, "skills", "excel-mcp", "SKILL.md"));
            var builtMcpSkill = File.ReadAllText(Path.Combine(outputDir, "excel-mcp", "skills", "excel-mcp", "SKILL.md"));
            Assert.Equal(sourceMcpSkill, builtMcpSkill);

            var sourceCliSkill = File.ReadAllText(Path.Combine(RepoRoot, "skills", "excel-cli", "SKILL.md"));
            var builtCliSkill = File.ReadAllText(Path.Combine(outputDir, "excel-cli", "skills", "excel-cli", "SKILL.md"));
            Assert.Equal(sourceCliSkill, builtCliSkill);

            AssertSkillDirectoryMatchesSource(
                Path.Combine(RepoRoot, "skills", "excel-mcp"),
                Path.Combine(outputDir, "excel-mcp", "skills", "excel-mcp"),
                version);
            AssertSkillDirectoryMatchesSource(
                Path.Combine(RepoRoot, "skills", "excel-cli"),
                Path.Combine(outputDir, "excel-cli", "skills", "excel-cli"),
                version);
            AssertLocalSkillLinksResolve(Path.Combine(outputDir, "excel-mcp", "skills", "excel-mcp"));
            AssertLocalSkillLinksResolve(Path.Combine(outputDir, "excel-cli", "skills", "excel-cli"));

            var overlayInstallGlobal = File.ReadAllText(Path.Combine(RepoRoot, ".github", "plugins", "excel-cli", "com.github.copilot", "bin", "install-global.ps1"));
            var builtInstallGlobal = File.ReadAllText(Path.Combine(outputDir, "excel-cli", "com.github.copilot", "bin", "install-global.ps1"));
            Assert.Equal(overlayInstallGlobal, builtInstallGlobal);

            var overlayCliBootstrap = File.ReadAllText(Path.Combine(RepoRoot, ".github", "plugins", "excel-cli", "bin", "start-cli.ps1"));
            var builtCliBootstrap = File.ReadAllText(Path.Combine(outputDir, "excel-cli", "bin", "start-cli.ps1"));
            Assert.Equal(overlayCliBootstrap, builtCliBootstrap);

            var overlayCliDownload = File.ReadAllText(Path.Combine(RepoRoot, ".github", "plugins", "excel-cli", "bin", "download.ps1"));
            var builtCliDownload = File.ReadAllText(Path.Combine(outputDir, "excel-cli", "bin", "download.ps1"));
            Assert.Equal(overlayCliDownload, builtCliDownload);

            var overlayMcpBootstrap = File.ReadAllText(Path.Combine(RepoRoot, ".github", "plugins", "excel-mcp", "bin", "start-mcp.ps1"));
            var builtMcpBootstrap = File.ReadAllText(Path.Combine(outputDir, "excel-mcp", "bin", "start-mcp.ps1"));
            Assert.Equal(overlayMcpBootstrap, builtMcpBootstrap);

            var overlayMcpInstallGlobal = File.ReadAllText(Path.Combine(RepoRoot, ".github", "plugins", "excel-mcp", "com.github.copilot", "bin", "install-global.ps1"));
            var builtMcpInstallGlobal = File.ReadAllText(Path.Combine(outputDir, "excel-mcp", "com.github.copilot", "bin", "install-global.ps1"));
            Assert.Equal(overlayMcpInstallGlobal, builtMcpInstallGlobal);

            var overlayMcpDownload = File.ReadAllText(Path.Combine(RepoRoot, ".github", "plugins", "excel-mcp", "bin", "download.ps1"));
            var builtMcpDownload = File.ReadAllText(Path.Combine(outputDir, "excel-mcp", "bin", "download.ps1"));
            Assert.Equal(overlayMcpDownload, builtMcpDownload);
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task BuildPlugins_IncludesCliCommandAndSharedReferencesInExcelCliSkill()
    {
        var sandbox = CreateSandbox("build-includes-cli-command-reference");
        try
        {
            var outputDir = Path.Combine(sandbox, "built-plugins");
            var version = "9.9.13-test";

            var result = await RunPowerShellFileAsync(
                BuildPluginsScript,
                [
                    "-Version", version,
                    "-OutputDir", outputDir
                ]);

            Assert.Equal(0, result.ExitCode);

            var sourceReferencePath = Path.Combine(RepoRoot, "skills", "excel-cli", "references", "cli-commands.md");
            var builtReferencePath = Path.Combine(outputDir, "excel-cli", "skills", "excel-cli", "references", "cli-commands.md");
            var sourceSharedReferences = Directory.GetFiles(Path.Combine(RepoRoot, "skills", "shared"), "*.md");

            Assert.True(File.Exists(builtReferencePath), $"Expected excel-cli plugin to package CLI command reference at {builtReferencePath}");
            Assert.Equal(File.ReadAllText(sourceReferencePath), File.ReadAllText(builtReferencePath));
            foreach (var sourceSharedReference in sourceSharedReferences)
            {
                var builtSharedReference = Path.Combine(
                    outputDir,
                    "excel-cli",
                    "skills",
                    "excel-cli",
                    "references",
                    Path.GetFileName(sourceSharedReference));
                Assert.True(File.Exists(builtSharedReference), $"Expected excel-cli plugin to package shared reference at {builtSharedReference}");
                var builtContent = File.ReadAllText(builtSharedReference);
                Assert.Contains("CLI syntax note", builtContent);
                Assert.Contains(
                    NormalizeLineEndings(File.ReadAllText(sourceSharedReference)),
                    NormalizeLineEndings(builtContent));

                var builtMcpReference = Path.Combine(
                    outputDir,
                    "excel-mcp",
                    "skills",
                    "excel-mcp",
                    "references",
                    Path.GetFileName(sourceSharedReference));
                Assert.True(File.Exists(builtMcpReference), $"Expected excel-mcp plugin to package shared reference at {builtMcpReference}");
                Assert.Equal(File.ReadAllText(sourceSharedReference), File.ReadAllText(builtMcpReference));
            }
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task BuildPlugins_UsesSourceOwnedTemplatesWithoutPublishedRepository()
    {
        var sandbox = CreateSandbox("build-source-owned-templates");
        try
        {
            var outputDir = Path.Combine(sandbox, "built-plugins");

            var result = await RunPowerShellFileAsync(
                BuildPluginsScript,
                [
                    "-Version", "9.9.11-test",
                    "-OutputDir", outputDir
                ]);

            Assert.Equal(0, result.ExitCode);
            AssertAgentPluginManifest(Path.Combine(outputDir, "excel-mcp"), "9.9.11-test");
            AssertAgentPluginManifest(Path.Combine(outputDir, "excel-cli"), "9.9.11-test");
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task BuildPlugins_SmokeRun_ExitsZeroAndPrintsAsciiSummary()
    {
        var sandbox = CreateSandbox("build-smoke-summary");
        try
        {
            var outputDir = Path.Combine(sandbox, "built-plugins");
            var version = "9.9.12-test";

            var result = await RunPowerShellFileAsync(
                BuildPluginsScript,
                [
                    "-Version", version,
                    "-OutputDir", outputDir
                ]);

            Assert.Equal(0, result.ExitCode);
            Assert.Contains("=== Build Complete ===", result.Stdout);
            Assert.Contains($"Version: {version}", result.Stdout);
            Assert.Contains($"Output:  {outputDir}", result.Stdout);
            Assert.Contains("[ok] excel-mcp - bootstrap assets and skill", result.Stdout);
            Assert.Contains("[ok] excel-cli - bootstrap assets and skill", result.Stdout);
            Assert.Contains($@"copilot plugin install {outputDir}\excel-mcp", result.Stdout);
            Assert.Contains($@"copilot plugin install {outputDir}\excel-cli", result.Stdout);
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task SyncPublishedPluginRepo_WritesCanonicalManifestAndCopiesBootstrapPlugins()
    {
        var sandbox = CreateSandbox("sync-published-plugin-repo");
        try
        {
            var builtPluginsDir = Path.Combine(sandbox, "built-plugins");
            var publishedRepoDir = Path.Combine(sandbox, "published-repo");
            var version = "9.9.12-test";

            Directory.CreateDirectory(publishedRepoDir);
            File.WriteAllText(Path.Combine(publishedRepoDir, "marketplace.json"), "{}");

            var buildResult = await RunPowerShellFileAsync(
                BuildPluginsScript,
                [
                    "-Version", version,
                    "-OutputDir", builtPluginsDir
                ]);

            Assert.Equal(0, buildResult.ExitCode);

            var syncResult = await RunPowerShellFileAsync(
                SyncPublishedRepoScript,
                [
                    "-PublishedRepoDir", publishedRepoDir,
                    "-BuiltPluginsDir", builtPluginsDir,
                    "-Version", version
                ]);

            Assert.Equal(0, syncResult.ExitCode);

            var canonicalManifestPath = Path.Combine(publishedRepoDir, ".github", "plugin", "marketplace.json");
            Assert.True(File.Exists(canonicalManifestPath), $"Canonical marketplace manifest should exist at {canonicalManifestPath}");
            Assert.False(File.Exists(Path.Combine(publishedRepoDir, "marketplace.json")));

            using var manifest = JsonDocument.Parse(File.ReadAllText(canonicalManifestPath));
            var plugins = manifest.RootElement.GetProperty("plugins");
            Assert.Equal(2, plugins.GetArrayLength());

            var pluginEntries = plugins.EnumerateArray().ToDictionary(
                p => p.GetProperty("name").GetString()!,
                p => p);

            Assert.Equal("./plugins/excel-mcp", pluginEntries["excel-mcp"].GetProperty("source").GetString());
            Assert.Equal("./plugins/excel-cli", pluginEntries["excel-cli"].GetProperty("source").GetString());
            Assert.Contains("./plugins/excel-mcp/skills/excel-mcp", pluginEntries["excel-mcp"].GetProperty("skills").EnumerateArray().Select(s => s.GetString()));
            Assert.Contains("./plugins/excel-cli/skills/excel-cli", pluginEntries["excel-cli"].GetProperty("skills").EnumerateArray().Select(s => s.GetString()));
            AssertMarketplacePluginMetadata(pluginEntries["excel-mcp"]);
            AssertMarketplacePluginMetadata(pluginEntries["excel-cli"]);

            AssertBootstrapAssetSet(
                Path.Combine(publishedRepoDir, "plugins", "excel-mcp"),
                "mcp.json",
                "bin/start-mcp.ps1",
                "bin/download.ps1",
                "com.github.copilot/bin/install-global.ps1");

            AssertBootstrapAssetSet(
                Path.Combine(publishedRepoDir, "plugins", "excel-cli"),
                "bin/start-cli.ps1",
                "bin/download.ps1",
                "com.github.copilot/bin/install-global.ps1");

            AssertBootstrapAssetsAbsent(
                Path.Combine(publishedRepoDir, "plugins", "excel-mcp"),
                ".mcp.json",
                "bin/download-mcp.ps1",
                "bin/bootstrap-state.json");

            AssertBootstrapAssetsAbsent(
                Path.Combine(publishedRepoDir, "plugins", "excel-cli"),
                "bin/download-cli.ps1",
                "bin/bootstrap-state.json");

            var publishedValidationScript = Path.Combine(publishedRepoDir, "tests", "Test-Plugins.ps1");
            var sourceValidationScript = Path.Combine(
                RepoRoot,
                ".github",
                "plugins",
                "marketplace-repo",
                "tests",
                "Test-Plugins.ps1");
            Assert.True(File.Exists(publishedValidationScript));
            Assert.Equal(
                File.ReadAllText(sourceValidationScript),
                File.ReadAllText(publishedValidationScript));

            var publishedInstructionsPath = Path.Combine(
                publishedRepoDir,
                ".github",
                "copilot-instructions.md");
            var sourceInstructionsPath = Path.Combine(
                RepoRoot,
                ".github",
                "plugins",
                "marketplace-repo",
                ".github",
                "copilot-instructions.md");
            Assert.True(File.Exists(sourceInstructionsPath));
            Assert.True(File.Exists(publishedInstructionsPath));
            Assert.Equal(
                File.ReadAllText(sourceInstructionsPath),
                File.ReadAllText(publishedInstructionsPath));
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Theory]
    [InlineData("excel-cli", "excelcli.exe", "ExcelMcp-CLI-{0}-windows.zip")]
    [InlineData("excel-mcp", "mcp-excel.exe", "ExcelMcp-MCP-Server-{0}-windows.zip")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task DownloadBootstrap_FirstRun_AutoDownloadsLatestWindowsRuntime(
        string pluginName,
        string executableName,
        string assetNameFormat)
    {
        Assert.True(OperatingSystem.IsWindows(), "Plugin bootstrap packaging tests require Windows.");

        var sandbox = CreateSandbox($"download-first-run-{pluginName}");
        try
        {
            var userProfile = Path.Combine(sandbox, "user");
            Directory.CreateDirectory(userProfile);

            var harnessPath = CreateDownloadHarnessScript(sandbox);
            var version = "1.2.3";
            var tag = $"v{version}";
            var assetName = string.Format(CultureInfo.InvariantCulture, assetNameFormat, version);

            var result = await RunPowerShellFileAsync(
                harnessPath,
                [
                    "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                    "-ExecutableName", executableName,
                    "-Tag", tag,
                    "-AssetName", assetName,
                    "-Mode", "success"
                ],
                environmentVariables: new Dictionary<string, string>
                {
                    ["USERPROFILE"] = userProfile,
                    ["COPILOT_AGENT_SESSION_ID"] = "session-a",
                    ["OS"] = "Windows_NT"
                });

            Assert.Equal(0, result.ExitCode);

            var statePath = GetBootstrapStatePath(userProfile, pluginName);
            Assert.True(File.Exists(statePath), $"Expected bootstrap state at {statePath}");

            using var state = JsonDocument.Parse(File.ReadAllText(statePath));
            Assert.Equal(tag, state.RootElement.GetProperty("latestTag").GetString());
            Assert.Equal(version, state.RootElement.GetProperty("latestVersion").GetString());
            Assert.Equal(assetName, state.RootElement.GetProperty("assetName").GetString());
            Assert.Matches(
                "^[0-9a-f]{64}$",
                state.RootElement.GetProperty("expectedSha256").GetString()!);

            var binaryPath = state.RootElement.GetProperty("binaryPath").GetString();
            Assert.False(string.IsNullOrWhiteSpace(binaryPath));
            Assert.True(File.Exists(binaryPath!), $"Expected resolved runtime at {binaryPath}");
            Assert.EndsWith(executableName, binaryPath, StringComparison.OrdinalIgnoreCase);

            Assert.Equal(1, ReadMockCallCount(userProfile, "rest"));
            Assert.Equal(1, ReadMockCallCount(userProfile, "checksum"));
            Assert.Equal(1, ReadMockCallCount(userProfile, "web"));
            Assert.Equal(1, ReadMockCallCount(userProfile, "expand"));
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Theory]
    [InlineData("excel-cli", "excelcli.exe", "ExcelMcp-CLI-{0}-windows.zip")]
    [InlineData("excel-mcp", "mcp-excel.exe", "ExcelMcp-MCP-Server-{0}-windows.zip")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task DownloadBootstrap_MissingChecksumAsset_FailsClosed(
        string pluginName,
        string executableName,
        string assetNameFormat)
    {
        Assert.True(OperatingSystem.IsWindows(), "Plugin bootstrap packaging tests require Windows.");

        var sandbox = CreateSandbox($"download-missing-checksum-{pluginName}");
        try
        {
            var userProfile = Path.Combine(sandbox, "user");
            Directory.CreateDirectory(userProfile);
            var version = "2.0.0";
            var assetName = string.Format(CultureInfo.InvariantCulture, assetNameFormat, version);

            var result = await RunPowerShellFileAsync(
                CreateDownloadHarnessScript(sandbox),
                [
                    "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                    "-ExecutableName", executableName,
                    "-Tag", $"v{version}",
                    "-AssetName", assetName,
                    "-Mode", "missing-checksum-asset"
                ],
                environmentVariables: new Dictionary<string, string>
                {
                    ["USERPROFILE"] = userProfile,
                    ["COPILOT_AGENT_SESSION_ID"] = "session-a",
                    ["OS"] = "Windows_NT"
                });

            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("SHA256SUMS", result.CombinedOutput, StringComparison.Ordinal);
            Assert.Contains("does not contain", result.CombinedOutput, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, ReadMockCallCount(userProfile, "web"));
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Theory]
    [InlineData("excel-cli", "excelcli.exe", "ExcelMcp-CLI-{0}-windows.zip", "malformed-checksum", "malformed")]
    [InlineData("excel-cli", "excelcli.exe", "ExcelMcp-CLI-{0}-windows.zip", "missing-checksum-entry", "does not contain")]
    [InlineData("excel-mcp", "mcp-excel.exe", "ExcelMcp-MCP-Server-{0}-windows.zip", "malformed-checksum", "malformed")]
    [InlineData("excel-mcp", "mcp-excel.exe", "ExcelMcp-MCP-Server-{0}-windows.zip", "missing-checksum-entry", "does not contain")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task DownloadBootstrap_InvalidChecksumMetadata_FailsClosed(
        string pluginName,
        string executableName,
        string assetNameFormat,
        string mode,
        string expectedMessage)
    {
        Assert.True(OperatingSystem.IsWindows(), "Plugin bootstrap packaging tests require Windows.");

        var sandbox = CreateSandbox($"download-invalid-checksum-{pluginName}-{mode}");
        try
        {
            var userProfile = Path.Combine(sandbox, "user");
            Directory.CreateDirectory(userProfile);
            var version = "2.0.0";
            var assetName = string.Format(CultureInfo.InvariantCulture, assetNameFormat, version);

            var result = await RunPowerShellFileAsync(
                CreateDownloadHarnessScript(sandbox),
                [
                    "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                    "-ExecutableName", executableName,
                    "-Tag", $"v{version}",
                    "-AssetName", assetName,
                    "-Mode", mode
                ],
                environmentVariables: new Dictionary<string, string>
                {
                    ["USERPROFILE"] = userProfile,
                    ["COPILOT_AGENT_SESSION_ID"] = "session-a",
                    ["OS"] = "Windows_NT"
                });

            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("SHA256SUMS", result.CombinedOutput, StringComparison.Ordinal);
            Assert.Contains(expectedMessage, result.CombinedOutput, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, ReadMockCallCount(userProfile, "web"));
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Theory]
    [InlineData("excel-cli", "excelcli.exe", "ExcelMcp-CLI-{0}-windows.zip")]
    [InlineData("excel-mcp", "mcp-excel.exe", "ExcelMcp-MCP-Server-{0}-windows.zip")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task DownloadBootstrap_TamperedDownload_FailsClosed(
        string pluginName,
        string executableName,
        string assetNameFormat)
    {
        Assert.True(OperatingSystem.IsWindows(), "Plugin bootstrap packaging tests require Windows.");

        var sandbox = CreateSandbox($"download-tampered-{pluginName}");
        try
        {
            var userProfile = Path.Combine(sandbox, "user");
            Directory.CreateDirectory(userProfile);
            var version = "2.0.0";
            var assetName = string.Format(CultureInfo.InvariantCulture, assetNameFormat, version);

            var result = await RunPowerShellFileAsync(
                CreateDownloadHarnessScript(sandbox),
                [
                    "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                    "-ExecutableName", executableName,
                    "-Tag", $"v{version}",
                    "-AssetName", assetName,
                    "-Mode", "checksum-mismatch"
                ],
                environmentVariables: new Dictionary<string, string>
                {
                    ["USERPROFILE"] = userProfile,
                    ["COPILOT_AGENT_SESSION_ID"] = "session-a",
                    ["OS"] = "Windows_NT"
                });

            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("SHA-256 mismatch", result.CombinedOutput, StringComparison.Ordinal);
            Assert.Equal(2, ReadMockCallCount(userProfile, "web"));
            Assert.Equal(0, ReadMockCallCount(userProfile, "expand"));
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Theory]
    [InlineData("excel-cli", "excelcli.exe", "ExcelMcp-CLI-{0}-windows.zip")]
    [InlineData("excel-mcp", "mcp-excel.exe", "ExcelMcp-MCP-Server-{0}-windows.zip")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task DownloadBootstrap_PluginHost_UsesPluginDataRuntime(
        string pluginName,
        string executableName,
        string assetNameFormat)
    {
        Assert.True(OperatingSystem.IsWindows(), "Plugin bootstrap packaging tests require Windows.");

        var sandbox = CreateSandbox($"download-plugin-data-{pluginName}");
        try
        {
            var userProfile = Path.Combine(sandbox, "user");
            var pluginData = Path.Combine(sandbox, "plugin-data");
            Directory.CreateDirectory(userProfile);
            Directory.CreateDirectory(pluginData);

            var harnessPath = CreateDownloadHarnessScript(sandbox);
            var version = "1.2.3";
            var tag = $"v{version}";
            var assetName = string.Format(CultureInfo.InvariantCulture, assetNameFormat, version);

            var result = await RunPowerShellFileAsync(
                harnessPath,
                [
                    "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                    "-ExecutableName", executableName,
                    "-Tag", tag,
                    "-AssetName", assetName,
                    "-Mode", "success"
                ],
                environmentVariables: new Dictionary<string, string>
                {
                    ["USERPROFILE"] = userProfile,
                    ["PLUGIN_DATA"] = pluginData,
                    ["COPILOT_AGENT_SESSION_ID"] = "session-a",
                    ["OS"] = "Windows_NT"
                });

            Assert.Equal(0, result.ExitCode);

            var pluginDataStatePath = Path.Combine(pluginData, "runtime", "bootstrap-state.json");
            Assert.True(File.Exists(pluginDataStatePath), $"Expected plugin bootstrap state at {pluginDataStatePath}");
            Assert.False(File.Exists(GetBootstrapStatePath(userProfile, pluginName)));

            using var state = JsonDocument.Parse(File.ReadAllText(pluginDataStatePath));
            var binaryPath = state.RootElement.GetProperty("binaryPath").GetString();
            Assert.False(string.IsNullOrWhiteSpace(binaryPath));
            Assert.StartsWith(
                Path.Combine(pluginData, "runtime"),
                binaryPath!,
                StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Theory]
    [InlineData("excel-cli", "excelcli.exe", "ExcelMcp-CLI-{0}-windows.zip")]
    [InlineData("excel-mcp", "mcp-excel.exe", "ExcelMcp-MCP-Server-{0}-windows.zip")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task DownloadBootstrap_SameSession_DoesNotRecheckGitHubRelease(
        string pluginName,
        string executableName,
        string assetNameFormat)
    {
        Assert.True(OperatingSystem.IsWindows(), "Plugin bootstrap packaging tests require Windows.");

        var sandbox = CreateSandbox($"download-same-session-{pluginName}");
        try
        {
            var userProfile = Path.Combine(sandbox, "user");
            Directory.CreateDirectory(userProfile);

            var harnessPath = CreateDownloadHarnessScript(sandbox);
            var version = "1.2.3";
            var tag = $"v{version}";
            var assetName = string.Format(CultureInfo.InvariantCulture, assetNameFormat, version);
            var env = new Dictionary<string, string>
            {
                ["USERPROFILE"] = userProfile,
                ["COPILOT_AGENT_SESSION_ID"] = "session-a",
                ["OS"] = "Windows_NT"
            };

            var firstResult = await RunPowerShellFileAsync(
                harnessPath,
                [
                    "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                    "-ExecutableName", executableName,
                    "-Tag", tag,
                    "-AssetName", assetName,
                    "-Mode", "success"
                ],
                environmentVariables: env);

            Assert.Equal(0, firstResult.ExitCode);

            ResetMockCalls(userProfile);

            var secondResult = await RunPowerShellFileAsync(
                harnessPath,
                [
                    "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                    "-ExecutableName", executableName,
                    "-Tag", tag,
                    "-AssetName", assetName,
                    "-Mode", "api-fail"
                ],
                environmentVariables: env);

            Assert.Equal(0, secondResult.ExitCode);
            Assert.Contains("Freshness already checked for this Copilot session.", secondResult.CombinedOutput);
            Assert.Equal(0, ReadMockCallCount(userProfile, "rest"));
            Assert.Equal(0, ReadMockCallCount(userProfile, "web"));
            Assert.Equal(0, ReadMockCallCount(userProfile, "expand"));
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Theory]
    [InlineData("excel-cli", "excelcli.exe", "ExcelMcp-CLI-{0}-windows.zip")]
    [InlineData("excel-mcp", "mcp-excel.exe", "ExcelMcp-MCP-Server-{0}-windows.zip")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task DownloadBootstrap_NewRelease_RefreshesStaleRuntime(
        string pluginName,
        string executableName,
        string assetNameFormat)
    {
        Assert.True(OperatingSystem.IsWindows(), "Plugin bootstrap packaging tests require Windows.");

        var sandbox = CreateSandbox($"download-stale-refresh-{pluginName}");
        try
        {
            var userProfile = Path.Combine(sandbox, "user");
            Directory.CreateDirectory(userProfile);

            var harnessPath = CreateDownloadHarnessScript(sandbox);
            var env = new Dictionary<string, string>
            {
                ["USERPROFILE"] = userProfile,
                ["OS"] = "Windows_NT"
            };

            var initialVersion = "1.2.3";
            var refreshedVersion = "1.2.4";

            env["COPILOT_AGENT_SESSION_ID"] = "session-a";
            var firstResult = await RunPowerShellFileAsync(
                harnessPath,
                [
                    "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                    "-ExecutableName", executableName,
                    "-Tag", $"v{initialVersion}",
                    "-AssetName", string.Format(CultureInfo.InvariantCulture, assetNameFormat, initialVersion),
                    "-Mode", "success"
                ],
                environmentVariables: env);

            Assert.Equal(0, firstResult.ExitCode);

            ResetMockCalls(userProfile);
            env["COPILOT_AGENT_SESSION_ID"] = "session-b";

            var secondResult = await RunPowerShellFileAsync(
                harnessPath,
                [
                    "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                    "-ExecutableName", executableName,
                    "-Tag", $"v{refreshedVersion}",
                    "-AssetName", string.Format(CultureInfo.InvariantCulture, assetNameFormat, refreshedVersion),
                    "-Mode", "success"
                ],
                environmentVariables: env);

            Assert.Equal(0, secondResult.ExitCode);

            var statePath = GetBootstrapStatePath(userProfile, pluginName);
            using var state = JsonDocument.Parse(File.ReadAllText(statePath));
            Assert.Equal($"v{refreshedVersion}", state.RootElement.GetProperty("latestTag").GetString());
            Assert.Equal(refreshedVersion, state.RootElement.GetProperty("latestVersion").GetString());

            var binaryPath = state.RootElement.GetProperty("binaryPath").GetString();
            Assert.False(string.IsNullOrWhiteSpace(binaryPath));
            Assert.Contains($@"releases\{refreshedVersion}\", binaryPath!, StringComparison.OrdinalIgnoreCase);

            Assert.Equal(1, ReadMockCallCount(userProfile, "rest"));
            Assert.Equal(1, ReadMockCallCount(userProfile, "web"));
            Assert.Equal(1, ReadMockCallCount(userProfile, "expand"));
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Theory]
    [InlineData("excel-cli", "excelcli.exe", "ExcelMcp-CLI-{0}-windows.zip", "Failed to resolve the latest excelcli release.")]
    [InlineData("excel-mcp", "mcp-excel.exe", "ExcelMcp-MCP-Server-{0}-windows.zip", "Failed to resolve the latest ExcelMcp MCP server release.")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task DownloadBootstrap_MissingWindowsAsset_SurfacesClearFailureMessage(
        string pluginName,
        string executableName,
        string assetNameFormat,
        string expectedErrorPrefix)
    {
        Assert.True(OperatingSystem.IsWindows(), "Plugin bootstrap packaging tests require Windows.");

        var sandbox = CreateSandbox($"download-failure-{pluginName}");
        try
        {
            var userProfile = Path.Combine(sandbox, "user");
            Directory.CreateDirectory(userProfile);

            var harnessPath = CreateDownloadHarnessScript(sandbox);
            var version = "2.0.0";
            var tag = $"v{version}";
            var assetName = string.Format(CultureInfo.InvariantCulture, assetNameFormat, version);

            var result = await RunPowerShellFileAsync(
                harnessPath,
                [
                    "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                    "-ExecutableName", executableName,
                    "-Tag", tag,
                    "-AssetName", assetName,
                    "-Mode", "missing-asset"
                ],
                environmentVariables: new Dictionary<string, string>
                {
                    ["USERPROFILE"] = userProfile,
                    ["COPILOT_AGENT_SESSION_ID"] = "session-a",
                    ["OS"] = "Windows_NT"
                });

            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains(expectedErrorPrefix, result.CombinedOutput);
            Assert.Contains(assetName, result.CombinedOutput);
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Theory]
    [InlineData("excel-cli", "excelcli.exe", "ExcelMcp-CLI-{0}-windows.zip")]
    [InlineData("excel-mcp", "mcp-excel.exe", "ExcelMcp-MCP-Server-{0}-windows.zip")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task DownloadBootstrap_ApiUnavailableWithWarmCache_FallsBackToCachedRuntime(
        string pluginName,
        string executableName,
        string assetNameFormat)
    {
        Assert.True(OperatingSystem.IsWindows(), "Plugin bootstrap packaging tests require Windows.");

        var sandbox = CreateSandbox($"download-offline-fallback-{pluginName}");
        try
        {
            var userProfile = Path.Combine(sandbox, "user");
            Directory.CreateDirectory(userProfile);

            var harnessPath = CreateDownloadHarnessScript(sandbox);
            var version = "1.2.3";
            var tag = $"v{version}";
            var assetName = string.Format(CultureInfo.InvariantCulture, assetNameFormat, version);

            var firstResult = await RunPowerShellFileAsync(
                harnessPath,
                [
                    "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                    "-ExecutableName", executableName,
                    "-Tag", tag,
                    "-AssetName", assetName,
                    "-Mode", "success"
                ],
                environmentVariables: new Dictionary<string, string>
                {
                    ["USERPROFILE"] = userProfile,
                    ["COPILOT_AGENT_SESSION_ID"] = "session-a",
                    ["OS"] = "Windows_NT"
                });

            Assert.Equal(0, firstResult.ExitCode);
            ResetMockCalls(userProfile);

            // A brand new Copilot session re-runs the freshness check, and that check now fails.
            var offlineResult = await RunPowerShellFileAsync(
                harnessPath,
                [
                    "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                    "-ExecutableName", executableName,
                    "-Tag", tag,
                    "-AssetName", assetName,
                    "-Mode", "api-fail",
                    "-QuietMode"
                ],
                environmentVariables: new Dictionary<string, string>
                {
                    ["USERPROFILE"] = userProfile,
                    ["COPILOT_AGENT_SESSION_ID"] = "session-b",
                    ["OS"] = "Windows_NT"
                });

            Assert.Equal(0, offlineResult.ExitCode);
            Assert.EndsWith(executableName, offlineResult.Stdout.Trim(), StringComparison.OrdinalIgnoreCase);

            // The warning must never reach stdout: it carries the resolved path, and for the MCP
            // plugin it is also the MCP stdio transport.
            Assert.Contains("Could not check for updates", offlineResult.Stderr, StringComparison.Ordinal);
            Assert.DoesNotContain("Could not check for updates", offlineResult.Stdout, StringComparison.Ordinal);

            // Falling back must not re-download anything.
            Assert.Equal(0, ReadMockCallCount(userProfile, "web"));
            Assert.Equal(0, ReadMockCallCount(userProfile, "expand"));
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Theory]
    [InlineData("excel-cli", "excelcli.exe", "ExcelMcp-CLI-{0}-windows.zip")]
    [InlineData("excel-mcp", "mcp-excel.exe", "ExcelMcp-MCP-Server-{0}-windows.zip")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task DownloadBootstrap_ChecksumUnavailableWithVerifiedWarmCache_FallsBackToCachedRuntime(
        string pluginName,
        string executableName,
        string assetNameFormat)
    {
        Assert.True(OperatingSystem.IsWindows(), "Plugin bootstrap packaging tests require Windows.");

        var sandbox = CreateSandbox($"download-checksum-offline-fallback-{pluginName}");
        try
        {
            var userProfile = Path.Combine(sandbox, "user");
            Directory.CreateDirectory(userProfile);
            var harnessPath = CreateDownloadHarnessScript(sandbox);
            var version = "1.2.3";
            var assetName = string.Format(CultureInfo.InvariantCulture, assetNameFormat, version);
            var env = new Dictionary<string, string>
            {
                ["USERPROFILE"] = userProfile,
                ["COPILOT_AGENT_SESSION_ID"] = "session-a",
                ["OS"] = "Windows_NT"
            };

            var firstResult = await RunPowerShellFileAsync(
                harnessPath,
                [
                    "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                    "-ExecutableName", executableName,
                    "-Tag", $"v{version}",
                    "-AssetName", assetName,
                    "-Mode", "success"
                ],
                environmentVariables: env);

            Assert.Equal(0, firstResult.ExitCode);
            ResetMockCalls(userProfile);
            env["COPILOT_AGENT_SESSION_ID"] = "session-b";

            var offlineResult = await RunPowerShellFileAsync(
                harnessPath,
                [
                    "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                    "-ExecutableName", executableName,
                    "-Tag", $"v{version}",
                    "-AssetName", assetName,
                    "-Mode", "checksum-download-fail",
                    "-QuietMode"
                ],
                environmentVariables: env);

            Assert.Equal(0, offlineResult.ExitCode);
            Assert.EndsWith(executableName, offlineResult.Stdout.Trim(), StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Could not check for updates", offlineResult.Stderr, StringComparison.Ordinal);
            Assert.Equal(1, ReadMockCallCount(userProfile, "checksum"));
            Assert.Equal(0, ReadMockCallCount(userProfile, "web"));
            Assert.Equal(0, ReadMockCallCount(userProfile, "expand"));
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Theory]
    [InlineData("excel-cli", "excelcli.exe", "ExcelMcp-CLI-{0}-windows.zip")]
    [InlineData("excel-mcp", "mcp-excel.exe", "ExcelMcp-MCP-Server-{0}-windows.zip")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task DownloadBootstrap_CachedArchiveRemoved_ReusesExtractedRuntimeWithoutDownloading(
        string pluginName,
        string executableName,
        string assetNameFormat)
    {
        Assert.True(OperatingSystem.IsWindows(), "Plugin bootstrap packaging tests require Windows.");

        var sandbox = CreateSandbox($"download-archive-removed-{pluginName}");
        try
        {
            var userProfile = Path.Combine(sandbox, "user");
            Directory.CreateDirectory(userProfile);

            var harnessPath = CreateDownloadHarnessScript(sandbox);
            var version = "1.2.3";
            var tag = $"v{version}";
            var assetName = string.Format(CultureInfo.InvariantCulture, assetNameFormat, version);

            var firstResult = await RunPowerShellFileAsync(
                harnessPath,
                [
                    "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                    "-ExecutableName", executableName,
                    "-Tag", tag,
                    "-AssetName", assetName,
                    "-Mode", "success"
                ],
                environmentVariables: new Dictionary<string, string>
                {
                    ["USERPROFILE"] = userProfile,
                    ["COPILOT_AGENT_SESSION_ID"] = "session-a",
                    ["OS"] = "Windows_NT"
                });

            Assert.Equal(0, firstResult.ExitCode);

            // Simulate a disk cleanup tool reclaiming the cached archive while the extracted
            // runtime stays in place.
            var downloadsDir = Path.Combine(
                userProfile,
                ".copilot",
                "plugin-runtime",
                "mcp-server-excel",
                pluginName,
                "downloads");
            Assert.True(Directory.Exists(downloadsDir), $"Expected cached downloads at {downloadsDir}");
            Directory.Delete(downloadsDir, recursive: true);

            ResetMockCalls(userProfile);

            var secondResult = await RunPowerShellFileAsync(
                harnessPath,
                [
                    "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                    "-ExecutableName", executableName,
                    "-Tag", tag,
                    "-AssetName", assetName,
                    "-Mode", "success",
                    "-QuietMode"
                ],
                environmentVariables: new Dictionary<string, string>
                {
                    ["USERPROFILE"] = userProfile,
                    ["COPILOT_AGENT_SESSION_ID"] = "session-b",
                    ["OS"] = "Windows_NT"
                });

            Assert.Equal(0, secondResult.ExitCode);
            Assert.EndsWith(executableName, secondResult.Stdout.Trim(), StringComparison.OrdinalIgnoreCase);

            // The already-extracted runtime matches the resolved release, so nothing is fetched
            // or re-extracted just because the archive is gone.
            Assert.Equal(0, ReadMockCallCount(userProfile, "web"));
            Assert.Equal(0, ReadMockCallCount(userProfile, "expand"));
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Theory]
    [InlineData("excel-cli", "excelcli.exe", "ExcelMcp-CLI-{0}-windows.zip", "Failed to resolve the latest excelcli release.")]
    [InlineData("excel-mcp", "mcp-excel.exe", "ExcelMcp-MCP-Server-{0}-windows.zip", "Failed to resolve the latest ExcelMcp MCP server release.")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task DownloadBootstrap_ApiUnavailableWithoutCache_FailsWithClearError(
        string pluginName,
        string executableName,
        string assetNameFormat,
        string expectedErrorPrefix)
    {
        Assert.True(OperatingSystem.IsWindows(), "Plugin bootstrap packaging tests require Windows.");

        var sandbox = CreateSandbox($"download-offline-cold-{pluginName}");
        try
        {
            var userProfile = Path.Combine(sandbox, "user");
            Directory.CreateDirectory(userProfile);

            var harnessPath = CreateDownloadHarnessScript(sandbox);
            var version = "1.2.3";
            var tag = $"v{version}";
            var assetName = string.Format(CultureInfo.InvariantCulture, assetNameFormat, version);

            var result = await RunPowerShellFileAsync(
                harnessPath,
                [
                    "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                    "-ExecutableName", executableName,
                    "-Tag", tag,
                    "-AssetName", assetName,
                    "-Mode", "api-fail"
                ],
                environmentVariables: new Dictionary<string, string>
                {
                    ["USERPROFILE"] = userProfile,
                    ["COPILOT_AGENT_SESSION_ID"] = "session-a",
                    ["OS"] = "Windows_NT"
                });

            // With no cached runtime to fall back to there is nothing to degrade to, so the
            // failure must stay loud rather than emitting an unusable path.
            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains(expectedErrorPrefix, result.CombinedOutput, StringComparison.Ordinal);
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Theory]
    [InlineData("excel-cli", "excelcli.exe", "ExcelMcp-CLI-{0}-windows.zip")]
    [InlineData("excel-mcp", "mcp-excel.exe", "ExcelMcp-MCP-Server-{0}-windows.zip")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task DownloadBootstrap_CorruptCachedArchive_SelfHeals(
        string pluginName,
        string executableName,
        string assetNameFormat)
    {
        Assert.True(OperatingSystem.IsWindows(), "Plugin bootstrap packaging tests require Windows.");

        var sandbox = CreateSandbox($"download-corrupt-zip-{pluginName}");
        try
        {
            var userProfile = Path.Combine(sandbox, "user");
            Directory.CreateDirectory(userProfile);

            var harnessPath = CreateDownloadHarnessScript(sandbox);
            var version = "1.2.3";
            var tag = $"v{version}";
            var assetName = string.Format(CultureInfo.InvariantCulture, assetNameFormat, version);

            var firstResult = await RunPowerShellFileAsync(
                harnessPath,
                [
                    "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                    "-ExecutableName", executableName,
                    "-Tag", tag,
                    "-AssetName", assetName,
                    "-Mode", "success"
                ],
                environmentVariables: new Dictionary<string, string>
                {
                    ["USERPROFILE"] = userProfile,
                    ["COPILOT_AGENT_SESSION_ID"] = "session-a",
                    ["OS"] = "Windows_NT"
                });

            Assert.Equal(0, firstResult.ExitCode);

            // Replace the cached archive with a different but still readable ZIP. Archive probing
            // alone accepts it, so only the published SHA-256 can detect the tampering.
            var pluginCache = Path.Combine(userProfile, ".copilot", "plugin-runtime", "mcp-server-excel", pluginName);
            var tamperedDirectory = Path.Combine(sandbox, "tampered");
            Directory.CreateDirectory(tamperedDirectory);
            File.WriteAllText(Path.Combine(tamperedDirectory, executableName), "tampered runtime");
            var cachedArchivePath = Path.Combine(pluginCache, "downloads", assetName);
            File.Delete(cachedArchivePath);
            ZipFile.CreateFromDirectory(tamperedDirectory, cachedArchivePath);
            Directory.Delete(Path.Combine(pluginCache, "releases", version), recursive: true);

            ResetMockCalls(userProfile);

            var secondResult = await RunPowerShellFileAsync(
                harnessPath,
                [
                    "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                    "-ExecutableName", executableName,
                    "-Tag", tag,
                    "-AssetName", assetName,
                    "-Mode", "success",
                    "-QuietMode"
                ],
                environmentVariables: new Dictionary<string, string>
                {
                    ["USERPROFILE"] = userProfile,
                    ["COPILOT_AGENT_SESSION_ID"] = "session-b",
                    ["OS"] = "Windows_NT"
                });

            Assert.Equal(0, secondResult.ExitCode);
            Assert.EndsWith(executableName, secondResult.Stdout.Trim(), StringComparison.OrdinalIgnoreCase);
            Assert.True(File.Exists(secondResult.Stdout.Trim()), "Expected a usable runtime after self-healing.");

            // The tampered archive must be discarded and refetched rather than reused.
            Assert.Equal(1, ReadMockCallCount(userProfile, "web"));

            // Hash validation happens before extraction, so only the verified replacement is read.
            Assert.Equal(1, ReadMockCallCount(userProfile, "expand"));
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Theory]
    [InlineData("excel-cli", "excelcli.exe", "ExcelMcp-CLI-{0}-windows.zip")]
    [InlineData("excel-mcp", "mcp-excel.exe", "ExcelMcp-MCP-Server-{0}-windows.zip")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task DownloadBootstrap_StandaloneInstall_ReChecksOnceTheWindowElapses(
        string pluginName,
        string executableName,
        string assetNameFormat)
    {
        Assert.True(OperatingSystem.IsWindows(), "Plugin bootstrap packaging tests require Windows.");

        var sandbox = CreateSandbox($"download-standalone-window-{pluginName}");
        try
        {
            var userProfile = Path.Combine(sandbox, "user");
            Directory.CreateDirectory(userProfile);

            var harnessPath = CreateDownloadHarnessScript(sandbox);
            var version = "1.2.3";
            var assetName = string.Format(CultureInfo.InvariantCulture, assetNameFormat, version);

            // No COPILOT_AGENT_SESSION_ID: this is the PATH/shim install, where the session id is
            // the constant "standalone" and therefore always equals the previously recorded one.
            var env = new Dictionary<string, string>
            {
                ["USERPROFILE"] = userProfile,
                ["OS"] = "Windows_NT"
            };

            var firstResult = await RunPowerShellFileAsync(
                harnessPath,
                [
                    "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                    "-ExecutableName", executableName,
                    "-Tag", $"v{version}",
                    "-AssetName", assetName,
                    "-Mode", "success"
                ],
                environmentVariables: env);

            Assert.Equal(0, firstResult.ExitCode);

            ResetMockCalls(userProfile);

            var immediateResult = await RunPowerShellFileAsync(
                harnessPath,
                [
                    "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                    "-ExecutableName", executableName,
                    "-Tag", $"v{version}",
                    "-AssetName", assetName,
                    "-Mode", "success",
                    "-QuietMode"
                ],
                environmentVariables: env);

            Assert.Equal(0, immediateResult.ExitCode);
            Assert.Equal(0, ReadMockCallCount(userProfile, "rest"));

            // Age the recorded check past the staleness window.
            var statePath = GetBootstrapStatePath(userProfile, pluginName);
            var state = JsonNode.Parse(File.ReadAllText(statePath))!;
            state["checkedAtUtc"] = DateTime.UtcNow.AddHours(-48).ToString("o", CultureInfo.InvariantCulture);
            File.WriteAllText(statePath, state.ToJsonString());

            ResetMockCalls(userProfile);

            var agedResult = await RunPowerShellFileAsync(
                harnessPath,
                [
                    "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                    "-ExecutableName", executableName,
                    "-Tag", $"v{version}",
                    "-AssetName", assetName,
                    "-Mode", "success",
                    "-QuietMode"
                ],
                environmentVariables: env);

            Assert.Equal(0, agedResult.ExitCode);
            Assert.Equal(1, ReadMockCallCount(userProfile, "rest"));
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Theory]
    [InlineData("excel-cli", "excelcli.exe", "ExcelMcp-CLI-{0}-windows.zip", "GITHUB_TOKEN")]
    [InlineData("excel-mcp", "mcp-excel.exe", "ExcelMcp-MCP-Server-{0}-windows.zip", "GH_TOKEN")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task DownloadBootstrap_TokenInEnvironment_AuthenticatesReleaseLookup(
        string pluginName,
        string executableName,
        string assetNameFormat,
        string tokenVariable)
    {
        Assert.True(OperatingSystem.IsWindows(), "Plugin bootstrap packaging tests require Windows.");

        var sandbox = CreateSandbox($"download-token-{pluginName}");
        try
        {
            var userProfile = Path.Combine(sandbox, "user");
            Directory.CreateDirectory(userProfile);

            var harnessPath = CreateDownloadHarnessScript(sandbox);
            var version = "1.2.3";
            var assetName = string.Format(CultureInfo.InvariantCulture, assetNameFormat, version);

            var arguments = new[]
            {
                "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                "-ExecutableName", executableName,
                "-Tag", $"v{version}",
                "-AssetName", assetName,
                "-Mode", "success",
                "-QuietMode"
            };

            var anonymousResult = await RunPowerShellFileAsync(
                harnessPath,
                arguments,
                environmentVariables: new Dictionary<string, string>
                {
                    ["USERPROFILE"] = userProfile,
                    ["COPILOT_AGENT_SESSION_ID"] = "session-a",
                    ["OS"] = "Windows_NT"
                });

            Assert.Equal(0, anonymousResult.ExitCode);

            var headersPath = Path.Combine(userProfile, "mock-calls", "rest-headers.txt");
            Assert.DoesNotContain("Authorization=", File.ReadAllText(headersPath), StringComparison.Ordinal);

            var authenticatedResult = await RunPowerShellFileAsync(
                harnessPath,
                arguments,
                environmentVariables: new Dictionary<string, string>
                {
                    ["USERPROFILE"] = userProfile,
                    ["COPILOT_AGENT_SESSION_ID"] = "session-b",
                    ["OS"] = "Windows_NT",
                    [tokenVariable] = "token-value"
                });

            Assert.Equal(0, authenticatedResult.ExitCode);
            Assert.Contains("Authorization=Bearer token-value", File.ReadAllText(headersPath), StringComparison.Ordinal);
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Theory]
    [InlineData("excel-cli", "excelcli.exe", "ExcelMcp-CLI-{0}-windows.zip")]
    [InlineData("excel-mcp", "mcp-excel.exe", "ExcelMcp-MCP-Server-{0}-windows.zip")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task DownloadBootstrap_RunningRuntime_KeepsWorkingInstall(
        string pluginName,
        string executableName,
        string assetNameFormat)
    {
        Assert.True(OperatingSystem.IsWindows(), "Plugin bootstrap packaging tests require Windows.");

        var sandbox = CreateSandbox($"download-locked-runtime-{pluginName}");
        try
        {
            var userProfile = Path.Combine(sandbox, "user");
            Directory.CreateDirectory(userProfile);

            var harnessPath = CreateDownloadHarnessScript(sandbox);
            var version = "1.2.3";
            var assetName = string.Format(CultureInfo.InvariantCulture, assetNameFormat, version);
            var env = new Dictionary<string, string>
            {
                ["USERPROFILE"] = userProfile,
                ["COPILOT_AGENT_SESSION_ID"] = "session-a",
                ["OS"] = "Windows_NT"
            };

            var firstResult = await RunPowerShellFileAsync(
                harnessPath,
                [
                    "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                    "-ExecutableName", executableName,
                    "-Tag", $"v{version}",
                    "-AssetName", assetName,
                    "-Mode", "success"
                ],
                environmentVariables: env);

            Assert.Equal(0, firstResult.ExitCode);

            var releaseDir = Path.Combine(
                userProfile,
                ".copilot",
                "plugin-runtime",
                "mcp-server-excel",
                pluginName,
                "releases",
                version);

            var installedBinary = Directory.GetFiles(releaseDir, executableName, SearchOption.AllDirectories).Single();
            var companionFile = Path.Combine(Path.GetDirectoryName(installedBinary)!, "LICENSE.txt");
            File.WriteAllText(companionFile, "license text");

            env["COPILOT_AGENT_SESSION_ID"] = "session-b";

            // Hold the executable open, exactly as a running runtime would. Deleting the release
            // directory would remove every sibling file before failing on the locked executable,
            // which is how a working install used to end up half destroyed.
            using (File.Open(installedBinary, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                var lockedResult = await RunPowerShellFileAsync(
                    harnessPath,
                    [
                        "-ScriptPath", GetPluginScriptPath(pluginName, "download.ps1"),
                        "-ExecutableName", executableName,
                        "-Tag", $"v{version}",
                        "-AssetName", assetName,
                        "-Mode", "success",
                        "-QuietMode",
                        "-ForceMode"
                    ],
                    environmentVariables: env);

                Assert.Equal(0, lockedResult.ExitCode);
                Assert.Equal(installedBinary, lockedResult.Stdout.Trim(), ignoreCase: true);
            }

            Assert.True(File.Exists(installedBinary), "The in-use runtime must survive.");
            Assert.True(File.Exists(companionFile), "Sibling files must not be destroyed by a blocked upgrade.");
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task DownloadBootstrap_VersionCheck_ReadsStampedFileMetadata()
    {
        Assert.True(OperatingSystem.IsWindows(), "Plugin bootstrap packaging tests require Windows.");

        // Running the runtime with --version would trigger its own network update check, which is
        // exactly wrong inside a bootstrap that must work offline. The script therefore reads the
        // version stamped into the file, and this asserts that it agrees with the OS.
        var stampedBinary = Path.Combine(Environment.SystemDirectory, "notepad.exe");
        Assert.True(File.Exists(stampedBinary), $"Expected a version-stamped binary at {stampedBinary}");

        var expectedVersion = FileVersionInfo.GetVersionInfo(stampedBinary).ProductVersion!.Split('+')[0].Trim();

        var sandbox = CreateSandbox("download-version-metadata");
        try
        {
            var unstampedBinary = Path.Combine(sandbox, "unstamped.exe");
            File.WriteAllText(unstampedBinary, "not a real binary");

            var probePath = Path.Combine(sandbox, "probe.ps1");
            File.WriteAllText(
                probePath,
                $$"""
                $ErrorActionPreference = "Stop"
                Set-StrictMode -Version Latest

                $scriptText = [System.IO.File]::ReadAllText("{{GetPluginScriptPath("excel-cli", "download.ps1").Replace("\\", "\\\\")}}", [System.Text.UTF8Encoding]::new($false))
                $ast = [System.Management.Automation.Language.Parser]::ParseInput($scriptText, [ref]$null, [ref]$null)
                foreach ($name in @("Get-BinaryProductVersion", "Test-BinaryMatchesVersion")) {
                    $definition = $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true) | Select-Object -First 1
                    if ($null -eq $definition) { throw "download.ps1 no longer defines $name." }
                    . ([scriptblock]::Create($definition.Extent.Text))
                }

                Write-Output "stamped=$(Get-BinaryProductVersion -Path '{{stampedBinary.Replace("\\", "\\\\")}}')"
                Write-Output "unstamped-is-null=$([string]::IsNullOrWhiteSpace((Get-BinaryProductVersion -Path '{{unstampedBinary.Replace("\\", "\\\\")}}')))"
                Write-Output "matches=$(Test-BinaryMatchesVersion -Path '{{stampedBinary.Replace("\\", "\\\\")}}' -ExpectedVersion '{{expectedVersion}}')"
                Write-Output "mismatch=$(Test-BinaryMatchesVersion -Path '{{stampedBinary.Replace("\\", "\\\\")}}' -ExpectedVersion '0.0.1')"
                Write-Output "unstamped-accepted=$(Test-BinaryMatchesVersion -Path '{{unstampedBinary.Replace("\\", "\\\\")}}' -ExpectedVersion '0.0.1')"
                """);

            var result = await RunPowerShellFileAsync(probePath, []);

            Assert.Equal(0, result.ExitCode);
            Assert.Contains($"stamped={expectedVersion}", result.Stdout, StringComparison.Ordinal);

            // A file with no version resource cannot be identified, and an unidentifiable runtime
            // is still better than no runtime, so it is accepted rather than rejected.
            Assert.Contains("unstamped-is-null=True", result.Stdout, StringComparison.Ordinal);
            Assert.Contains("unstamped-accepted=True", result.Stdout, StringComparison.Ordinal);

            Assert.Contains("matches=True", result.Stdout, StringComparison.Ordinal);
            Assert.Contains("mismatch=False", result.Stdout, StringComparison.Ordinal);
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public void DownloadBootstrapScripts_AreStructurallyIdenticalAcrossPlugins()
    {
        // The two bootstrap scripts are maintained as copies that differ only by plugin-specific
        // names. Normalizing those tokens must make them byte-identical, so a fix applied to one
        // copy and forgotten in the other fails here instead of shipping.
        var cli = File.ReadAllText(GetPluginScriptPath("excel-cli", "download.ps1"));
        var mcp = File.ReadAllText(GetPluginScriptPath("excel-mcp", "download.ps1"));

        Assert.Equal(NormalizeBootstrapScript(cli), NormalizeBootstrapScript(mcp));
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public void CanonicalPluginDocumentation_DoesNotUseRetiredNamesOrFlags()
    {
        var retiredDocumentation = new (string Pattern, string Description)[]
        {
            (Regex.Escape("ExcelMcp-CLI-latest-windows.zip"), "nonexistent unversioned CLI release asset"),
            (Regex.Escape("excel-mcp-server.exe"), "retired MCP executable name"),
            (Regex.Escape("excel-mcp-bundle.mcpb"), "retired MCPB asset name"),
            (Regex.Escape("file(action: 'open', filePath"), "retired MCP file path parameter"),
            (Regex.Escape("file(action: 'close', sessionId"), "retired MCP session parameter"),
            (@"(?<![A-Za-z0-9-])--range-address(?![A-Za-z0-9-])", "retired CLI range flag"),
            (@"(?<![A-Za-z0-9-])--sheet-name(?![A-Za-z0-9-])", "retired CLI worksheet flag"),
            (@"(?<![A-Za-z0-9-])--source-table-name(?![A-Za-z0-9-])", "retired CLI PivotTable source flag")
        };
        var documentationRoots = new[]
        {
            Path.Combine(RepoRoot, ".github", "plugins", "excel-cli"),
            Path.Combine(RepoRoot, ".github", "plugins", "excel-mcp"),
            Path.Combine(RepoRoot, "docs", "guides"),
            Path.Combine(RepoRoot, "skills", "excel-mcp")
        };
        var failures = new List<string>();

        var documentationFiles = documentationRoots.SelectMany(
                root => Directory.GetFiles(root, "*.md", SearchOption.AllDirectories))
            .Append(Path.Combine(RepoRoot, ".github", "plugins", "marketplace-repo", "README.md"));

        foreach (var path in documentationFiles)
        {
            var content = File.ReadAllText(path);
            foreach (var retired in retiredDocumentation)
            {
                if (Regex.IsMatch(content, retired.Pattern))
                {
                    failures.Add($"{Path.GetRelativePath(RepoRoot, path)}: {retired.Description}");
                }
            }
        }

        Assert.True(
            failures.Count == 0,
            "Canonical plugin documentation contains retired names or flags:\n" +
            string.Join('\n', failures));
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task BuildBootstrapScripts_CheckPassesForCommittedCopies()
    {
        var scriptPath = Path.Combine(RepoRoot, "scripts", "Build-BootstrapScripts.ps1");
        var result = await RunPowerShellFileAsync(scriptPath, ["-Check"]);

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("matching", result.Stdout, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task BuildBootstrapScripts_CheckFailsWhenRenderedCopyDrifts()
    {
        var sandbox = CreateSandbox("bootstrap-render-drift");
        try
        {
            var scriptPath = Path.Combine(RepoRoot, "scripts", "Build-BootstrapScripts.ps1");
            var outputRoot = Path.Combine(sandbox, "generated");

            var generateResult = await RunPowerShellFileAsync(scriptPath, ["-OutputRoot", outputRoot]);
            Assert.Equal(0, generateResult.ExitCode);

            var driftedScript = Path.Combine(outputRoot, "excel-cli", "bin", "download.ps1");
            File.AppendAllText(driftedScript, "`r`n# drift test marker`r`n");

            var checkResult = await RunPowerShellFileAsync(scriptPath, ["-Check", "-OutputRoot", outputRoot]);

            Assert.NotEqual(0, checkResult.ExitCode);
            Assert.Contains("excel-cli", checkResult.Stdout + checkResult.Stderr, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("drift", checkResult.Stdout + checkResult.Stderr, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task BuildBootstrapScripts_GeneratesUtf8WithoutBomAndPreservesCheckmark()
    {
        var sandbox = CreateSandbox("bootstrap-render-bytes");
        try
        {
            var scriptPath = Path.Combine(RepoRoot, "scripts", "Build-BootstrapScripts.ps1");
            var outputRoot = Path.Combine(sandbox, "generated");

            var renderResult = await RunPowerShellFileAsync(scriptPath, ["-OutputRoot", outputRoot]);
            Assert.Equal(0, renderResult.ExitCode);

            var cliScript = Path.Combine(outputRoot, "excel-cli", "bin", "download.ps1");
            var bytes = File.ReadAllBytes(cliScript);

            Assert.DoesNotContain(new byte[] { 0xEF, 0xBB, 0xBF }, bytes.Take(3).ToArray());
            Assert.Contains(bytes, b => b == 0x0D && bytes[Array.IndexOf(bytes, b) + 1] == 0x0A);
            Assert.Contains("✅", File.ReadAllText(cliScript, Encoding.UTF8), StringComparison.Ordinal);
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    private static string NormalizeBootstrapScript(string text)
        => text
            .Replace("the latest excelcli release.", "the latest RUNTIME release.", StringComparison.Ordinal)
            .Replace("the latest ExcelMcp MCP server release.", "the latest RUNTIME release.", StringComparison.Ordinal)
            .Replace("✅ excelcli runtime ready.", "✅ RUNTIME ready.", StringComparison.Ordinal)
            .Replace("✅ ExcelMcp MCP runtime ready.", "✅ RUNTIME ready.", StringComparison.Ordinal)
            .Replace("ExcelMcp-CLI-", "ASSET-", StringComparison.Ordinal)
            .Replace("ExcelMcp-MCP-Server-", "ASSET-", StringComparison.Ordinal)
            .Replace("excelcli.exe", "RUNTIME.exe", StringComparison.Ordinal)
            .Replace("mcp-excel.exe", "RUNTIME.exe", StringComparison.Ordinal)
            .Replace("excel-cli", "PLUGIN", StringComparison.Ordinal)
            .Replace("excel-mcp", "PLUGIN", StringComparison.Ordinal);

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task InstallGlobal_WhenDownloadScriptMissing_FailsBeforeWritingShims()
    {
        Assert.True(OperatingSystem.IsWindows(), "Plugin bootstrap packaging tests require Windows.");

        // The generated .cmd shim resolves the runtime by invoking bin\download.ps1, so the
        // installer depends on that script existing. Without an explicit guard it would happily
        // write shims that only fail later, at first use, with an opaque error. Validating the
        // dependency up front has to happen *before* any shim or PATH mutation, which is what
        // this test pins: a missing download.ps1 must abort with nothing written.
        var sandbox = CreateSandbox("install-global-missing-download");
        try
        {
            var pluginDir = Path.Combine(sandbox, "excel-cli");
            var pluginBinDir = Path.Combine(pluginDir, "bin");
            var installerDir = Path.Combine(pluginDir, "com.github.copilot", "bin");
            Directory.CreateDirectory(pluginBinDir);
            Directory.CreateDirectory(installerDir);

            // The wrapper is present; only download.ps1 is absent. That isolates the new guard
            // from the pre-existing wrapper check, so this test cannot pass for the wrong reason.
            File.WriteAllText(Path.Combine(pluginBinDir, "start-cli.ps1"), "exit 0");

            var installerPath = Path.Combine(installerDir, "install-global.ps1");
            File.Copy(
                Path.Combine(RepoRoot, ".github", "plugins", "excel-cli", "com.github.copilot", "bin", "install-global.ps1"),
                installerPath);

            // Redirect the profile so a regression that proceeds past the guard writes its shims
            // into the sandbox instead of the real ~/.copilot/bin.
            var fakeHome = Path.Combine(sandbox, "home");
            Directory.CreateDirectory(fakeHome);

            var result = await RunPowerShellFileAsync(
                installerPath,
                [],
                new Dictionary<string, string> { ["USERPROFILE"] = fakeHome });

            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("download.ps1", result.Stderr, StringComparison.OrdinalIgnoreCase);

            // Failing "cleanly" means no partial install: no shim directory, and no shims.
            var shimDir = Path.Combine(fakeHome, ".copilot", "bin");
            Assert.False(
                Directory.Exists(shimDir),
                $"Installer created {shimDir} despite the missing bootstrap script.");
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task StartCliWrapper_EscapesArgumentsSoTheyRoundTripThroughWin32Parsing()
    {
        Assert.True(OperatingSystem.IsWindows(), "Plugin bootstrap packaging tests require Windows.");

        var sandbox = CreateSandbox("start-cli-argument-fidelity");
        try
        {
            var harnessPath = Path.Combine(sandbox, "escape-harness.ps1");
            File.WriteAllText(harnessPath, """
                [CmdletBinding()]
                param([Parameter(Mandatory = $true)][string]$ScriptPath)

                $ErrorActionPreference = "Stop"

                $tokens = $null
                $errors = $null
                $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$errors)
                if ($errors.Count -gt 0) {
                    throw "Parse errors in $ScriptPath"
                }

                $definition = $ast.Find(
                    {
                        param($node)
                        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                        $node.Name -eq 'ConvertTo-NativeArgument'
                    },
                    $true)

                if ($null -eq $definition) {
                    throw "ConvertTo-NativeArgument was not found in $ScriptPath"
                }

                . ([scriptblock]::Create($definition.Extent.Text))

                $cases = @(
                    '[["Name","Amount"],["Widget",1500]]'
                    'plain'
                    'has space'
                    'trailing\'
                    'embedded\\backslash'
                    'quote"inside'
                    'backslash\"quote'
                    ''
                    'C:\reports\Q1 results.xlsx'
                    '{"nested":{"value":"a b"}}'
                )

                $encoded = foreach ($case in $cases) { ConvertTo-NativeArgument -Value $case }
                Write-Output ($encoded -join ' ')
                """);

            var result = await RunPowerShellFileAsync(
                harnessPath,
                ["-ScriptPath", GetPluginScriptPath("excel-cli", "start-cli.ps1")]);

            Assert.Equal(0, result.ExitCode);

            var commandLine = result.Stdout.Trim();
            Assert.NotEmpty(commandLine);

            string[] expected =
            [
                """[["Name","Amount"],["Widget",1500]]""",
                "plain",
                "has space",
                @"trailing\",
                @"embedded\\backslash",
                """quote"inside""",
                """backslash\"quote""",
                string.Empty,
                @"C:\reports\Q1 results.xlsx",
                """{"nested":{"value":"a b"}}"""
            ];

            // CommandLineToArgvW is the parser the CRT and .NET use to split a process command
            // line, so round-tripping through it proves the child sees the original arguments.
            var parsed = SplitCommandLine($"excelcli.exe {commandLine}").Skip(1).ToArray();

            Assert.Equal(expected, parsed);
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginBootstrap")]
    public async Task StartCliWrapper_RelaysOutputThroughPowerShellPipeline()
    {
        Assert.True(OperatingSystem.IsWindows(), "Plugin bootstrap packaging tests require Windows.");

        var sandbox = CreateSandbox("start-cli-pipeline-output");
        try
        {
            var pluginBinDirectory = Path.Combine(sandbox, "plugin", "bin");
            Directory.CreateDirectory(pluginBinDirectory);
            File.Copy(
                GetPluginScriptPath("excel-cli", "start-cli.ps1"),
                Path.Combine(pluginBinDirectory, "start-cli.ps1"));

            File.WriteAllText(
                Path.Combine(pluginBinDirectory, "download.ps1"),
                """
                [CmdletBinding()]
                param(
                    [switch]$PassThru,
                    [switch]$Quiet
                )

                Write-Output (Get-Command "cscript.exe").Source
                """,
                new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

            var echoScriptPath = Path.Combine(sandbox, "echo-streams.js");
            File.WriteAllText(
                echoScriptPath,
                """
                WScript.StdOut.Write("pipeline-ok");
                WScript.StdErr.Write("pipeline-error");
                """,
                new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

            var harnessPath = Path.Combine(sandbox, "invoke-wrapper.ps1");
            File.WriteAllText(
                harnessPath,
                """
                [CmdletBinding()]
                param(
                    [Parameter(Mandatory = $true)]
                    [string]$WrapperPath,

                    [Parameter(Mandatory = $true)]
                    [string]$EchoScriptPath
                )

                $captured = (& $WrapperPath //nologo $EchoScriptPath | Out-String).Trim()
                Write-Output "captured=$captured"
                """,
                new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

            var result = await RunPowerShellFileAsync(
                harnessPath,
                [
                    "-WrapperPath", Path.Combine(pluginBinDirectory, "start-cli.ps1"),
                    "-EchoScriptPath", echoScriptPath
                ]);

            Assert.Equal(0, result.ExitCode);
            Assert.Equal("captured=pipeline-ok", result.Stdout.Trim());
            Assert.Equal("pipeline-error", result.Stderr.Trim());
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [SupportedOSPlatform("windows")]
    private static string[] SplitCommandLine(string commandLine)
    {
        var argv = CommandLineToArgvW(commandLine, out var count);
        if (argv == IntPtr.Zero)
        {
            throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        try
        {
            var results = new string[count];
            for (var i = 0; i < count; i++)
            {
                var itemPtr = Marshal.ReadIntPtr(argv, i * IntPtr.Size);
                results[i] = Marshal.PtrToStringUni(itemPtr) ?? string.Empty;
            }

            return results;
        }
        finally
        {
            LocalFree(argv);
        }
    }

    [DllImport("shell32.dll", EntryPoint = "CommandLineToArgvW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr CommandLineToArgvW(string lpCmdLine, out int pNumArgs);

    [DllImport("kernel32.dll", EntryPoint = "LocalFree", SetLastError = true)]
    private static extern IntPtr LocalFree(IntPtr hMem);

    private static void AssertBootstrapAssetSet(string pluginRoot, params string[] relativePaths)
    {
        foreach (var relativePath in relativePaths)
        {
            var fullPath = Path.Combine(pluginRoot, relativePath);
            Assert.True(File.Exists(fullPath), $"Expected bootstrap asset at {fullPath}");
        }
    }

    private static void AssertBootstrapAssetsAbsent(string pluginRoot, params string[] relativePaths)
    {
        foreach (var relativePath in relativePaths)
        {
            var fullPath = Path.Combine(pluginRoot, relativePath);
            Assert.False(File.Exists(fullPath), $"Did not expect legacy bootstrap asset at {fullPath}");
        }
    }

    private static void AssertAgentPluginManifest(string pluginRoot, string expectedVersion)
    {
        var manifestPath = Path.Combine(pluginRoot, "plugin.json");
        Assert.True(File.Exists(manifestPath), $"Expected Agent Plugin manifest at {manifestPath}");

        using var document = JsonDocument.Parse(File.ReadAllText(manifestPath));
        var root = document.RootElement;
        var allowedProperties = new HashSet<string>(StringComparer.Ordinal)
        {
            "$schema",
            "name",
            "version",
            "description",
            "author",
            "homepage",
            "repository",
            "license",
            "keywords",
            "extensions"
        };

        Assert.Equal(JsonValueKind.Object, root.ValueKind);
        Assert.Equal(AgentPluginSchema, root.GetProperty("$schema").GetString());
        Assert.Equal(Path.GetFileName(pluginRoot), root.GetProperty("name").GetString());
        Assert.Equal(expectedVersion, root.GetProperty("version").GetString());
        Assert.Equal(JsonValueKind.String, root.GetProperty("description").ValueKind);
        Assert.Equal(JsonValueKind.String, root.GetProperty("repository").ValueKind);
        Assert.All(root.EnumerateObject(), property => Assert.Contains(property.Name, allowedProperties));
        Assert.DoesNotContain(root.EnumerateObject(), property =>
            property.Name is "displayName" or "publisher" or "skills" or "mcpServers");
    }

    private static void AssertPortableMcpConfiguration(string pluginRoot)
    {
        var mcpPath = Path.Combine(pluginRoot, "mcp.json");
        Assert.True(File.Exists(mcpPath), $"Expected portable MCP configuration at {mcpPath}");

        using var document = JsonDocument.Parse(File.ReadAllText(mcpPath));
        var root = document.RootElement;
        Assert.Equal(AgentPluginMcpSchema, root.GetProperty("$schema").GetString());
        Assert.Equal(2, root.EnumerateObject().Count());

        var server = root.GetProperty("mcpServers").GetProperty("excel-mcp");
        Assert.Equal("stdio", server.GetProperty("type").GetString());
        Assert.Equal("powershell", server.GetProperty("command").GetString());
        Assert.DoesNotContain(' ', server.GetProperty("command").GetString()!);

        var args = server.GetProperty("args").EnumerateArray().Select(arg => arg.GetString()).ToArray();
        Assert.Contains("${PLUGIN_ROOT}/bin/start-mcp.ps1", args);
        Assert.DoesNotContain(args, arg => arg?.Contains("{pluginDir}", StringComparison.Ordinal) == true);
    }

    private static void AssertAgentSkill(string skillRoot, string expectedName)
    {
        var skillPath = Path.Combine(skillRoot, "SKILL.md");
        Assert.True(File.Exists(skillPath), $"Expected Agent Skill at {skillPath}");

        var lines = File.ReadAllLines(skillPath);
        Assert.True(lines.Length > 3);
        Assert.Equal("---", lines[0].Trim());

        var closingDelimiter = Array.FindIndex(lines, 1, line => line.Trim() == "---");
        Assert.True(closingDelimiter > 1, $"{skillPath} must contain YAML frontmatter.");

        var nameLine = lines[1..closingDelimiter]
            .Single(line => line.StartsWith("name:", StringComparison.Ordinal));
        var name = nameLine["name:".Length..].Trim();
        Assert.Equal(expectedName, name);
        Assert.Matches("^(?!.*--)[a-z0-9](?:[a-z0-9-]*[a-z0-9])?$", name);
        Assert.InRange(name.Length, 1, 64);

        var descriptionIndex = Array.FindIndex(
            lines,
            1,
            closingDelimiter - 1,
            line => line.StartsWith("description:", StringComparison.Ordinal));
        Assert.True(descriptionIndex > 0, $"{skillPath} must declare a description.");

        var description = string.Join(
            " ",
            lines[(descriptionIndex + 1)..closingDelimiter]
                .TakeWhile(line => line.Length > 0 && char.IsWhiteSpace(line[0]))
                .Select(line => line.Trim()));
        Assert.InRange(description.Length, 1, 1024);
        Assert.Contains("Use when", description, StringComparison.OrdinalIgnoreCase);

        var allowedFields = new HashSet<string>(StringComparer.Ordinal)
        {
            "name",
            "description",
            "license",
            "compatibility",
            "metadata",
            "allowed-tools"
        };
        var frontmatterFields = lines[1..closingDelimiter]
            .Where(line => line.Length > 0 && !char.IsWhiteSpace(line[0]) && line.Contains(':'))
            .Select(line => line[..line.IndexOf(':')])
            .ToArray();
        Assert.All(frontmatterFields, field => Assert.Contains(field, allowedFields));

        var compatibilityLine = lines[1..closingDelimiter]
            .Single(line => line.StartsWith("compatibility:", StringComparison.Ordinal));
        var compatibility = compatibilityLine["compatibility:".Length..].Trim();
        Assert.InRange(compatibility.Length, 1, 500);
    }

    private static void AssertMarketplacePluginMetadata(JsonElement plugin)
    {
        var allowedProperties = new HashSet<string>(StringComparer.Ordinal)
        {
            "name",
            "source",
            "description",
            "version",
            "author",
            "homepage",
            "repository",
            "license",
            "keywords",
            "skills"
        };

        Assert.All(plugin.EnumerateObject(), property => Assert.Contains(property.Name, allowedProperties));
        Assert.Equal(JsonValueKind.String, plugin.GetProperty("repository").ValueKind);
        Assert.Equal(JsonValueKind.Array, plugin.GetProperty("skills").ValueKind);
    }

    private static void AssertSkillDirectoryMatchesSource(
        string sourceSkillRoot,
        string builtSkillRoot,
        string? expectedVersion = null)
    {
        var sourceFiles = Directory.GetFiles(sourceSkillRoot, "*", SearchOption.AllDirectories)
            .Select(path => Path.GetRelativePath(sourceSkillRoot, path))
            .Where(path => path != "VERSION")
            .OrderBy(path => path, StringComparer.Ordinal)
            .ToArray();
        var builtFiles = Directory.GetFiles(builtSkillRoot, "*", SearchOption.AllDirectories)
            .Select(path => Path.GetRelativePath(builtSkillRoot, path))
            .Where(path => path != "VERSION")
            .OrderBy(path => path, StringComparer.Ordinal)
            .ToArray();

        // VERSION is stamped by the build rather than copied from the canonical skill, so it is
        // compared separately instead of being required to exist in the source.
        Assert.Equal(sourceFiles, builtFiles);
        foreach (var relativePath in sourceFiles)
        {
            Assert.Equal(
                File.ReadAllText(Path.Combine(sourceSkillRoot, relativePath)),
                File.ReadAllText(Path.Combine(builtSkillRoot, relativePath)));
        }

        if (expectedVersion != null)
        {
            Assert.Equal(expectedVersion, File.ReadAllText(Path.Combine(builtSkillRoot, "VERSION")).Trim());
        }
    }

    private static void AssertLocalSkillLinksResolve(string skillRoot)
    {
        var skillPath = Path.Combine(skillRoot, "SKILL.md");
        var content = File.ReadAllText(skillPath);
        var localLinks = Regex.Matches(content, @"\]\((\./[^)#]+)(?:#[^)]+)?\)")
            .Select(match => match.Groups[1].Value)
            .Distinct(StringComparer.Ordinal)
            .ToArray();

        foreach (var localLink in localLinks)
        {
            var relativePath = localLink[2..].Replace('/', Path.DirectorySeparatorChar);
            var linkedPath = Path.Combine(skillRoot, relativePath);
            Assert.True(File.Exists(linkedPath), $"Local Agent Skill link '{localLink}' does not resolve from {skillPath}.");
        }
    }

    private static string CreateSandbox(string name)
    {
        var sandbox = Path.Combine(RepoRoot, "scratch", "plugin-bootstrap-test", $"{name}-{Guid.NewGuid():N}");
        Directory.CreateDirectory(sandbox);
        return sandbox;
    }

    private static string NormalizeLineEndings(string value)
    {
        return value.Replace("\r\n", "\n", StringComparison.Ordinal);
    }

    private static void DeleteDirectoryIfExists(string path)
    {
        if (Directory.Exists(path))
        {
            Directory.Delete(path, recursive: true);
        }
    }

    private static string CreateDownloadHarnessScript(string sandbox)
    {
        var harnessPath = Path.Combine(sandbox, "bootstrap-harness.ps1");
        File.WriteAllText(
            harnessPath,
            """
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [string]$ScriptPath,

                [Parameter(Mandatory = $true)]
                [string]$ExecutableName,

                [Parameter(Mandatory = $true)]
                [string]$Tag,

                [Parameter(Mandatory = $true)]
                [string]$AssetName,

                [Parameter(Mandatory = $true)]
                [string]$Mode,

                [switch]$QuietMode,

                [switch]$ForceMode
            )

            $ErrorActionPreference = "Stop"
            Set-StrictMode -Version Latest

            Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

            $callDir = Join-Path $env:USERPROFILE "mock-calls"
            New-Item -ItemType Directory -Path $callDir -Force | Out-Null
            $fixtureDirectory = Join-Path $env:USERPROFILE "fixture-runtime"
            $fixtureArchivePath = Join-Path $env:USERPROFILE "fixture-runtime.zip"
            New-Item -ItemType Directory -Path $fixtureDirectory -Force | Out-Null
            Set-Content -Path (Join-Path $fixtureDirectory $ExecutableName) -Value "fake runtime" -Encoding UTF8
            if (Test-Path $fixtureArchivePath) {
                Remove-Item -Path $fixtureArchivePath -Force
            }
            [System.IO.Compression.ZipFile]::CreateFromDirectory($fixtureDirectory, $fixtureArchivePath)
            $fixtureSha256 = (Get-FileHash -Path $fixtureArchivePath -Algorithm SHA256).Hash.ToLowerInvariant()

            function Add-MockCall {
                param([Parameter(Mandatory = $true)][string]$Name)

                $counterPath = Join-Path $callDir "$Name.count"
                $count = if (Test-Path $counterPath) { [int](Get-Content $counterPath -Raw) } else { 0 }
                Set-Content -Path $counterPath -Value ($count + 1) -Encoding UTF8
            }

            function Invoke-RestMethod {
                param(
                    [string]$Uri,
                    [hashtable]$Headers
                )

                Add-MockCall -Name "rest"

                if ($null -ne $Headers) {
                    $headerLines = foreach ($key in ($Headers.Keys | Sort-Object)) { "$key=$($Headers[$key])" }
                    Set-Content -Path (Join-Path $callDir "rest-headers.txt") -Value $headerLines -Encoding UTF8
                }

                switch ($Mode) {
                    "api-fail" {
                        throw "Simulated GitHub API failure"
                    }
                    "missing-asset" {
                        return [pscustomobject]@{
                            tag_name = $Tag
                            assets = @(
                                [pscustomobject]@{
                                    name = "notes.txt"
                                    browser_download_url = "https://example.test/notes.txt"
                                }
                            )
                        }
                    }
                    "missing-checksum-asset" {
                        return [pscustomobject]@{
                            tag_name = $Tag
                            assets = @(
                                [pscustomobject]@{
                                    name = $AssetName
                                    browser_download_url = "https://example.test/$AssetName"
                                }
                            )
                        }
                    }
                    default {
                        return [pscustomobject]@{
                            tag_name = $Tag
                            assets = @(
                                [pscustomobject]@{
                                    name = "notes.txt"
                                    browser_download_url = "https://example.test/notes.txt"
                                },
                                [pscustomobject]@{
                                    name = $AssetName
                                    browser_download_url = "https://example.test/$AssetName"
                                },
                                [pscustomobject]@{
                                    name = "SHA256SUMS"
                                    browser_download_url = "https://example.test/SHA256SUMS"
                                }
                            )
                        }
                    }
                }
            }

            function Invoke-WebRequest {
                param(
                    [string]$Uri,
                    [string]$OutFile
                )

                if ($Uri.EndsWith("/SHA256SUMS", [System.StringComparison]::Ordinal)) {
                    Add-MockCall -Name "checksum"
                    if ($Mode -eq "checksum-download-fail") {
                        throw "Simulated checksum download failure"
                    }

                    switch ($Mode) {
                        "malformed-checksum" {
                            Set-Content -Path $OutFile -Value "not a checksum manifest" -Encoding ASCII
                        }
                        "missing-checksum-entry" {
                            Set-Content -Path $OutFile -Value "$fixtureSha256  another-asset.zip" -Encoding ASCII
                        }
                        "checksum-mismatch" {
                            Set-Content -Path $OutFile -Value "$('0' * 64)  $AssetName" -Encoding ASCII
                        }
                        default {
                            Set-Content -Path $OutFile -Value "$fixtureSha256  $AssetName" -Encoding ASCII
                        }
                    }
                    return
                }

                Add-MockCall -Name "web"

                if ($Mode -eq "download-fail") {
                    throw "Simulated download failure"
                }

                New-Item -ItemType Directory -Path (Split-Path -Parent $OutFile) -Force | Out-Null

                if ($Mode -eq "corrupt-download") {
                    # A truncated transfer: the file exists but is not a readable archive.
                    Set-Content -Path $OutFile -Value "not a zip" -Encoding UTF8
                    return
                }

                if (Test-Path $OutFile) {
                    Remove-Item -Path $OutFile -Force
                }
                Copy-Item -Path $fixtureArchivePath -Destination $OutFile
            }

            function Expand-Archive {
                param(
                    [string]$Path,
                    [string]$DestinationPath,
                    [switch]$Force
                )

                Add-MockCall -Name "expand"

                # Faithful to the real cmdlet: extraction of an unreadable archive fails. Without
                # this the mock would happily "extract" a truncated download and hide exactly the
                # corruption handling these tests exist to cover.
                $probe = $null
                try {
                    $probe = [System.IO.Compression.ZipFile]::OpenRead($Path)
                } catch {
                    throw "Simulated extraction failure: '$Path' is not a readable archive."
                } finally {
                    if ($null -ne $probe) { $probe.Dispose() }
                }

                New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
                if ($Mode -eq "missing-binary-after-extract") {
                    Set-Content -Path (Join-Path $DestinationPath "README.txt") -Value "no runtime" -Encoding UTF8
                    return
                }

                Set-Content -Path (Join-Path $DestinationPath $ExecutableName) -Value "fake runtime" -Encoding UTF8
            }

            $env:OS = "Windows_NT"
            $scriptArgs = @{ PassThru = $true }
            if ($QuietMode) { $scriptArgs["Quiet"] = $true }
            if ($ForceMode) { $scriptArgs["Force"] = $true }

            & $ScriptPath @scriptArgs
            """);

        return harnessPath;
    }

    private static string GetPluginScriptPath(string pluginName, string fileName)
        => Path.Combine(RepoRoot, ".github", "plugins", pluginName, "bin", fileName);

    private static string GetBootstrapStatePath(string userProfile, string pluginName)
        => Path.Combine(userProfile, ".copilot", "plugin-runtime", "mcp-server-excel", pluginName, "bootstrap-state.json");

    private static int ReadMockCallCount(string userProfile, string counterName)
    {
        var counterPath = Path.Combine(userProfile, "mock-calls", $"{counterName}.count");
        return File.Exists(counterPath)
            ? int.Parse(File.ReadAllText(counterPath).Trim(), CultureInfo.InvariantCulture)
            : 0;
    }

    private static void ResetMockCalls(string userProfile)
    {
        var callDir = Path.Combine(userProfile, "mock-calls");
        if (Directory.Exists(callDir))
        {
            Directory.Delete(callDir, recursive: true);
        }
    }

    private static string FindRepoRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Sbroenne.ExcelMcp.sln")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("Could not locate repository root from test output directory.");
    }

    private static async Task<ProcessResult> RunPowerShellFileAsync(
        string scriptPath,
        IReadOnlyList<string> arguments,
        Dictionary<string, string>? environmentVariables = null,
        int timeoutMs = 30000)
    {
        var escapedScriptPath = scriptPath.Replace("'", "''");
        var escapedArguments = arguments
            .Select(argument => argument.Length > 0 && argument[0] == '-'
                ? argument
                : $"'{argument.Replace("'", "''")}'");
        var commandText = $"& '{escapedScriptPath}' {string.Join(" ", escapedArguments)}";

        var startInfo = new ProcessStartInfo
        {
            FileName = "pwsh",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            WorkingDirectory = RepoRoot
        };

        startInfo.ArgumentList.Add("-NoProfile");
        startInfo.ArgumentList.Add("-ExecutionPolicy");
        startInfo.ArgumentList.Add("Bypass");
        startInfo.ArgumentList.Add("-Command");
        startInfo.ArgumentList.Add(commandText);

        // The bootstrap reads ambient environment. Scrub the variables it consults so that a
        // developer machine or CI runner which happens to define them cannot silently change
        // what these tests exercise; callers opt back in by passing them explicitly.
        foreach (var ambientName in new[] { "COPILOT_AGENT_SESSION_ID", "GITHUB_TOKEN", "GH_TOKEN" })
        {
            startInfo.Environment.Remove(ambientName);
        }

        if (environmentVariables != null)
        {
            foreach (var (key, value) in environmentVariables)
            {
                startInfo.Environment[key] = value;
            }
        }

        using var process = new Process { StartInfo = startInfo };
        var stdout = new StringBuilder();
        var stderr = new StringBuilder();

        process.OutputDataReceived += (_, e) =>
        {
            if (e.Data != null)
            {
                stdout.AppendLine(e.Data);
            }
        };

        process.ErrorDataReceived += (_, e) =>
        {
            if (e.Data != null)
            {
                stderr.AppendLine(e.Data);
            }
        };

        process.Start();
        process.BeginOutputReadLine();
        process.BeginErrorReadLine();

        using var timeout = new CancellationTokenSource(timeoutMs);
        try
        {
            await process.WaitForExitAsync(timeout.Token);
        }
        catch (OperationCanceledException)
        {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException($"PowerShell script '{scriptPath}' timed out after {timeoutMs}ms.");
        }

        return new ProcessResult(process.ExitCode, stdout.ToString(), stderr.ToString());
    }

    private sealed record ProcessResult(int ExitCode, string Stdout, string Stderr)
    {
        public string CombinedOutput => $"{Stdout}{Environment.NewLine}{Stderr}";
    }
}
