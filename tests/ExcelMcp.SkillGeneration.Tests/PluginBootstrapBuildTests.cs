using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

/// <summary>
/// Integration tests for the Copilot plugin bootstrap packaging flow.
/// These exercise the real PowerShell build/sync scripts against synthetic plugin templates
/// so runtime bootstrap assets survive packaging without touching real user state.
/// </summary>
public sealed class PluginBootstrapBuildTests
{
    private const string AgentPluginSchema = "https://agent-plugins.org/schemas/1.0.0/plugin.schema.json";
    private const string AgentPluginMcpSchema = "https://agent-plugins.org/schemas/1.0.0/mcp.schema.json";
    private static readonly string RepoRoot = FindRepoRoot();
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
                Path.Combine(outputDir, "excel-cli", "skills", "excel-cli"));
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
    public async Task BuildPlugins_IncludesCliCommandReferenceInExcelCliSkillReferences()
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

            Assert.True(File.Exists(builtReferencePath), $"Expected excel-cli plugin to package CLI command reference at {builtReferencePath}");
            Assert.Equal(File.ReadAllText(sourceReferencePath), File.ReadAllText(builtReferencePath));
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

            var binaryPath = state.RootElement.GetProperty("binaryPath").GetString();
            Assert.False(string.IsNullOrWhiteSpace(binaryPath));
            Assert.True(File.Exists(binaryPath!), $"Expected resolved runtime at {binaryPath}");
            Assert.EndsWith(executableName, binaryPath, StringComparison.OrdinalIgnoreCase);

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
            .OrderBy(path => path, StringComparer.Ordinal)
            .ToArray();
        var builtFiles = Directory.GetFiles(builtSkillRoot, "*", SearchOption.AllDirectories)
            .Select(path => Path.GetRelativePath(builtSkillRoot, path))
            .OrderBy(path => path, StringComparer.Ordinal)
            .ToArray();

        Assert.Equal(sourceFiles, builtFiles);
        foreach (var relativePath in sourceFiles.Where(path => path != "VERSION"))
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

        Assert.NotEmpty(localLinks);
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
                [string]$Mode
            )

            $ErrorActionPreference = "Stop"
            Set-StrictMode -Version Latest

            $callDir = Join-Path $env:USERPROFILE "mock-calls"
            New-Item -ItemType Directory -Path $callDir -Force | Out-Null

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

                Add-MockCall -Name "web"

                if ($Mode -eq "download-fail") {
                    throw "Simulated download failure"
                }

                New-Item -ItemType Directory -Path (Split-Path -Parent $OutFile) -Force | Out-Null
                Set-Content -Path $OutFile -Value "fake zip payload" -Encoding UTF8
            }

            function Expand-Archive {
                param(
                    [string]$Path,
                    [string]$DestinationPath,
                    [switch]$Force
                )

                Add-MockCall -Name "expand"

                New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
                if ($Mode -eq "missing-binary-after-extract") {
                    Set-Content -Path (Join-Path $DestinationPath "README.txt") -Value "no runtime" -Encoding UTF8
                    return
                }

                Set-Content -Path (Join-Path $DestinationPath $ExecutableName) -Value "fake runtime" -Encoding UTF8
            }

            $env:OS = "Windows_NT"
            & $ScriptPath -PassThru
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
