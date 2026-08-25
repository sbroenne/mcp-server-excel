using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

/// <summary>
/// Integration tests for release metadata synchronization and workflow wiring.
/// </summary>
public sealed class ReleaseMetadataScriptTests
{
    private static readonly string RepoRoot = FindRepoRoot();
    private static readonly string UpdateMetadataScript = Path.Combine(
        RepoRoot,
        "scripts",
        "Update-McpRegistryMetadata.ps1");
    private static readonly string UpdateReleaseVersionScript = Path.Combine(
        RepoRoot,
        "scripts",
        "Update-ReleaseVersionMetadata.ps1");
    private static readonly string BuildChangelogScript = Path.Combine(
        RepoRoot,
        "scripts",
        "Build-Changelog.ps1");
    private static readonly string ReleaseWorkflow = Path.Combine(
        RepoRoot,
        ".github",
        "workflows",
        "release.yml");

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "ReleaseMetadata")]
    public async Task UpdateReleaseVersion_StampsEveryPersistentVersionSource()
    {
        var sandbox = CreateSandbox();
        try
        {
            CopyReleaseMetadataFiles(sandbox);
            var manifestPath = Path.Combine(sandbox, "mcpb", "manifest.json");
            var manifest = JsonNode.Parse(await File.ReadAllTextAsync(manifestPath))!.AsObject();
            manifest["nestedMetadata"] = new JsonObject
            {
                ["version"] = "nested-version"
            };
            await File.WriteAllTextAsync(
                manifestPath,
                manifest.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));

            var result = await RunPowerShellScriptAsync(
                UpdateReleaseVersionScript,
                ["-RepoRoot", sandbox, "-Version", "9.8.7"]);

            Assert.True(result.ExitCode == 0, result.CombinedOutput);
            AssertReleaseVersions(sandbox, "9.8.7");
            Assert.Equal(
                "nested-version",
                ReadJsonProperty(manifestPath, "nestedMetadata", "version"));
        }
        finally
        {
            Directory.Delete(sandbox, recursive: true);
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "ReleaseMetadata")]
    public void SourceTree_PersistentVersionsMatchCanonicalPackageVersion()
    {
        var expectedVersion = ReadJsonProperty(
            Path.Combine(RepoRoot, "package.json"),
            "version");

        AssertReleaseVersions(RepoRoot, expectedVersion);

        Assert.Equal(
            "0.0.0",
            ReadJsonProperty(
                Path.Combine(RepoRoot, ".github", "plugins", "excel-mcp", "plugin.json"),
                "version"));
        Assert.Equal(
            "0.0.0",
            ReadJsonProperty(
                Path.Combine(RepoRoot, ".github", "plugins", "excel-cli", "plugin.json"),
                "version"));
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "ReleaseMetadata")]
    public void ReleaseFlow_UsesAndCommitsCanonicalVersionUpdater()
    {
        var buildChangelog = File.ReadAllText(BuildChangelogScript);
        Assert.Contains(
            "& $updateReleaseVersionScript -RepoRoot $RepoRoot -Version $Version",
            buildChangelog,
            StringComparison.Ordinal);

        var releaseWorkflow = File.ReadAllText(ReleaseWorkflow);
        Assert.Contains(
            @".\scripts\Update-ReleaseVersionMetadata.ps1 -Version $env:VERSION",
            releaseWorkflow,
            StringComparison.Ordinal);
        Assert.DoesNotContain(
            "$content = $content -replace '<Version>",
            releaseWorkflow,
            StringComparison.Ordinal);

        var stagedPaths = new[]
        {
            "CHANGELOG.md",
            "package.json",
            "package-lock.json",
            "Directory.Build.props",
            "mcpb/manifest.json",
            "src/ExcelMcp.McpServer/.mcp/server.json",
            "vscode-extension/package.json",
            "vscode-extension/package-lock.json",
            ".changeset"
        };

        var stagingStart = releaseWorkflow.IndexOf("git add -A", StringComparison.Ordinal);
        Assert.True(stagingStart >= 0);
        var stagingEnd = releaseWorkflow.IndexOf(
            "if git diff --cached --quiet",
            stagingStart,
            StringComparison.Ordinal);
        Assert.True(stagingEnd > stagingStart);
        var stagingCommand = releaseWorkflow[stagingStart..stagingEnd];

        Assert.All(
            stagedPaths,
            path => Assert.Contains(path, stagingCommand, StringComparison.Ordinal));
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "ReleaseMetadata")]
    public async Task UpdateMetadata_StampsSeparatedTopLevelAndPackageVersions()
    {
        var sandbox = CreateSandbox();
        try
        {
            var metadataPath = Path.Combine(sandbox, "server.json");
            File.Copy(
                Path.Combine(RepoRoot, "src", "ExcelMcp.McpServer", ".mcp", "server.json"),
                metadataPath);

            var result = await RunPowerShellScriptAsync(
                UpdateMetadataScript,
                ["-ServerJsonPath", metadataPath, "-Version", "9.8.7"]);

            Assert.True(result.ExitCode == 0, result.CombinedOutput);
            using var document = JsonDocument.Parse(File.ReadAllText(metadataPath));
            var root = document.RootElement;
            Assert.Equal("9.8.7", root.GetProperty("version").GetString());

            var packages = root.GetProperty("packages").EnumerateArray().ToArray();
            var mcpPackage = Assert.Single(packages, package =>
                package.GetProperty("identifier").GetString() == "Sbroenne.ExcelMcp.McpServer");
            Assert.Equal("9.8.7", mcpPackage.GetProperty("version").GetString());
        }
        finally
        {
            Directory.Delete(sandbox, recursive: true);
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "ReleaseMetadata")]
    public async Task UpdateMetadata_RejectsMissingMcpServerPackage()
    {
        var sandbox = CreateSandbox();
        try
        {
            var metadataPath = Path.Combine(sandbox, "server.json");
            await File.WriteAllTextAsync(metadataPath, """{"version":"1.0.0","packages":[]}""");

            var result = await RunPowerShellScriptAsync(
                UpdateMetadataScript,
                ["-ServerJsonPath", metadataPath, "-Version", "9.8.7"]);

            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("exactly one", result.CombinedOutput, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            Directory.Delete(sandbox, recursive: true);
        }
    }

    private static string CreateSandbox()
    {
        var sandbox = Path.Combine(Path.GetTempPath(), $"ExcelMcpReleaseMetadata-{Guid.NewGuid():N}");
        Directory.CreateDirectory(sandbox);
        return sandbox;
    }

    private static void CopyReleaseMetadataFiles(string sandbox)
    {
        var relativePaths = new[]
        {
            "Directory.Build.props",
            "package.json",
            "package-lock.json",
            Path.Combine("mcpb", "manifest.json"),
            Path.Combine("src", "ExcelMcp.McpServer", ".mcp", "server.json"),
            Path.Combine("vscode-extension", "package.json"),
            Path.Combine("vscode-extension", "package-lock.json")
        };

        foreach (var relativePath in relativePaths)
        {
            var destinationPath = Path.Combine(sandbox, relativePath);
            Directory.CreateDirectory(Path.GetDirectoryName(destinationPath)!);
            File.Copy(Path.Combine(RepoRoot, relativePath), destinationPath);
        }
    }

    private static void AssertReleaseVersions(string root, string expectedVersion)
    {
        Assert.Equal(
            expectedVersion,
            ReadJsonProperty(Path.Combine(root, "package.json"), "version"));

        var props = XDocument.Load(Path.Combine(root, "Directory.Build.props"));
        var propertyGroup = Assert.Single(
            props.Root!.Elements("PropertyGroup"),
            group => group.Element("Version") != null);
        Assert.Equal(expectedVersion, propertyGroup.Element("Version")!.Value);
        Assert.Equal($"{expectedVersion}.0", propertyGroup.Element("AssemblyVersion")!.Value);
        Assert.Equal($"{expectedVersion}.0", propertyGroup.Element("FileVersion")!.Value);

        Assert.Equal(
            expectedVersion,
            ReadJsonProperty(Path.Combine(root, "mcpb", "manifest.json"), "version"));
        Assert.Equal(
            expectedVersion,
            ReadJsonProperty(
                Path.Combine(root, "src", "ExcelMcp.McpServer", ".mcp", "server.json"),
                "version"));
        Assert.Equal(
            expectedVersion,
            ReadJsonProperty(
                Path.Combine(root, "src", "ExcelMcp.McpServer", ".mcp", "server.json"),
                "packages",
                "0",
                "version"));
        Assert.Equal(
            expectedVersion,
            ReadJsonProperty(Path.Combine(root, "package-lock.json"), "version"));
        Assert.Equal(
            expectedVersion,
            ReadJsonProperty(Path.Combine(root, "package-lock.json"), "packages", "", "version"));
        Assert.Equal(
            expectedVersion,
            ReadJsonProperty(Path.Combine(root, "vscode-extension", "package.json"), "version"));
        Assert.Equal(
            expectedVersion,
            ReadJsonProperty(Path.Combine(root, "vscode-extension", "package-lock.json"), "version"));
        Assert.Equal(
            expectedVersion,
            ReadJsonProperty(
                Path.Combine(root, "vscode-extension", "package-lock.json"),
                "packages",
                "",
                "version"));
    }

    private static string ReadJsonProperty(string path, params string[] segments)
    {
        using var document = JsonDocument.Parse(File.ReadAllText(path));
        var element = document.RootElement;
        foreach (var segment in segments)
        {
            element = element.ValueKind == JsonValueKind.Array
                ? element[Convert.ToInt32(segment, System.Globalization.CultureInfo.InvariantCulture)]
                : element.GetProperty(segment);
        }

        return element.GetString()!;
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

    private static async Task<ScriptResult> RunPowerShellScriptAsync(
        string scriptPath,
        IReadOnlyList<string> arguments)
    {
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
        startInfo.ArgumentList.Add("-File");
        startInfo.ArgumentList.Add(scriptPath);
        foreach (var argument in arguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        using var process = Process.Start(startInfo);
        Assert.NotNull(process);

        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        await process.WaitForExitAsync(timeout.Token);

        return new ScriptResult(process.ExitCode, await stdout, await stderr);
    }

    private sealed record ScriptResult(int ExitCode, string Stdout, string Stderr)
    {
        public string CombinedOutput => $"{Stdout}{Environment.NewLine}{Stderr}";
    }
}
