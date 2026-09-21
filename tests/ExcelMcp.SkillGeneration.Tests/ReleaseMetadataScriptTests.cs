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
    public void ReleaseFlow_PackagesAndTagsCurrentChangelog()
    {
        var releaseWorkflow = File.ReadAllText(ReleaseWorkflow);
        var prepareRelease = ExtractWorkflowJob(releaseWorkflow, "prepare-release");
        var buildVsCode = ExtractWorkflowJob(releaseWorkflow, "build-vscode");
        var buildMcpb = ExtractWorkflowJob(releaseWorkflow, "build-mcpb");
        var createTag = ExtractWorkflowJob(releaseWorkflow, "create-tag");
        var createRelease = ExtractWorkflowJob(releaseWorkflow, "create-release");

        Assert.Contains("./scripts/Build-Changelog.ps1", prepareRelease, StringComparison.Ordinal);
        Assert.Contains("name: release-metadata", prepareRelease, StringComparison.Ordinal);
        Assert.Contains("needs: [version, prepare-release]", buildVsCode, StringComparison.Ordinal);
        Assert.Contains("name: release-metadata", buildVsCode, StringComparison.Ordinal);
        Assert.Contains(
            "Copy-Item \"prepared-release/CHANGELOG.md\" \"CHANGELOG.md\" -Force",
            buildVsCode,
            StringComparison.Ordinal);
        Assert.Contains("needs: [version, prepare-release]", buildMcpb, StringComparison.Ordinal);
        Assert.Contains("name: release-metadata", buildMcpb, StringComparison.Ordinal);
        Assert.Contains(
            "Copy-Item \"prepared-release/CHANGELOG.md\" \"CHANGELOG.md\" -Force",
            buildMcpb,
            StringComparison.Ordinal);

        var commitIndex = createTag.IndexOf("Commit Release Metadata Update", StringComparison.Ordinal);
        var tagIndex = createTag.IndexOf("Create and push tag", StringComparison.Ordinal);
        Assert.Contains("prepare-release", createTag, StringComparison.Ordinal);
        Assert.True(commitIndex >= 0, "The tag job must commit the generated release metadata.");
        Assert.True(tagIndex > commitIndex, "Release metadata must be committed before the tag is created.");
        Assert.Contains("git tag -a \"$TAG\" \"$RELEASE_COMMIT\"", createTag, StringComparison.Ordinal);
        Assert.Contains("SOURCE_SHA: ${{ github.sha }}", createTag, StringComparison.Ordinal);
        Assert.Contains("[ \"$BASE_SHA\" != \"$SOURCE_SHA\" ]", createTag, StringComparison.Ordinal);
        Assert.DoesNotContain("[skip ci]", createTag, StringComparison.Ordinal);

        Assert.Contains(
            "artifacts/release-metadata/release_notes_body.md",
            createRelease,
            StringComparison.Ordinal);
        Assert.DoesNotContain("./scripts/Build-Changelog.ps1", createRelease, StringComparison.Ordinal);
        Assert.DoesNotContain("Commit Release Metadata Update", createRelease, StringComparison.Ordinal);
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "ReleaseMetadata")]
    public void ReleaseFlow_EveryConsumerAppliesGeneratedCounts()
    {
        var releaseWorkflow = File.ReadAllText(ReleaseWorkflow);
        var jobs = ExtractWorkflowJobs(releaseWorkflow);

        var expectedConsumers = new[]
        {
            "build-cli",
            "build-mcp-server",
            "build-vscode",
            "build-mcpb",
            "build-agent-skills",
            "publish-mcp-registry",
            "create-tag"
        };

        Assert.Contains("release-doc-counts.patch", jobs["prepare-release"], StringComparison.Ordinal);

        var actualConsumers = jobs
            .Where(job => job.Key != "prepare-release"
                && job.Value.Contains("release-doc-counts.patch", StringComparison.Ordinal))
            .Select(job => job.Key)
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToArray();

        Assert.Equal(
            expectedConsumers.OrderBy(name => name, StringComparer.Ordinal).ToArray(),
            actualConsumers);

        foreach (var consumer in expectedConsumers)
        {
            var job = jobs[consumer];
            Assert.Contains("prepare-release", job, StringComparison.Ordinal);
            Assert.Contains("name: release-metadata", job, StringComparison.Ordinal);
            Assert.Contains("path: prepared-release", job, StringComparison.Ordinal);
            Assert.Contains(
                "git apply --whitespace=nowarn $patch.FullName",
                job,
                StringComparison.Ordinal);
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "ReleaseMetadata")]
    public async Task DocumentationCounts_UpdatePersistValidateAndRejectIncompatibleModes()
    {
        var sandbox = CreateSandbox();
        try
        {
            var sourceReadme = await File.ReadAllTextAsync(Path.Combine(RepoRoot, "README.md"));
            var headline = System.Text.RegularExpressions.Regex.Match(
                sourceReadme,
                @"(?<tools>\d+) tools with (?<operations>\d+) operations");
            Assert.True(headline.Success);
            var canonicalTools = int.Parse(headline.Groups["tools"].Value, System.Globalization.CultureInfo.InvariantCulture);
            var canonicalOperations = int.Parse(
                headline.Groups["operations"].Value,
                System.Globalization.CultureInfo.InvariantCulture);

            CopyDocumentationCountFiles(sandbox, canonicalTools, canonicalOperations);
            var readmePath = Path.Combine(sandbox, "README.md");
            var hooksPath = Path.Combine(sandbox, "gh-pages", "hooks.py");
            await File.WriteAllTextAsync(
                readmePath,
                (await File.ReadAllTextAsync(readmePath))
                    .Replace(
                        $"{canonicalTools} tools with {canonicalOperations} operations",
                        "1 tools with 2 operations",
                        StringComparison.Ordinal)
                    .Replace($"all {canonicalOperations} operations", "all 2 operations", StringComparison.Ordinal));

            var update = await RunPowerShellScriptAsync(
                Path.Combine(sandbox, "scripts", "check-doc-counts.ps1"),
                ["-Update", "-SkipBuild"],
                sandbox);

            Assert.True(update.ExitCode == 0, update.CombinedOutput);
            Assert.Contains(
                $"{canonicalTools} tools with {canonicalOperations} operations",
                await File.ReadAllTextAsync(readmePath),
                StringComparison.Ordinal);
            Assert.Contains(
                $"all {canonicalOperations} operations",
                await File.ReadAllTextAsync(readmePath),
                StringComparison.Ordinal);
            Assert.Contains(
                $"{canonicalTools} tools and {canonicalOperations} operations",
                await File.ReadAllTextAsync(hooksPath),
                StringComparison.Ordinal);

            var validation = await RunPowerShellScriptAsync(
                Path.Combine(sandbox, "scripts", "check-doc-counts.ps1"),
                ["-SkipBuild"],
                sandbox);
            Assert.True(validation.ExitCode == 0, validation.CombinedOutput);

            await File.WriteAllTextAsync(
                readmePath,
                (await File.ReadAllTextAsync(readmePath))
                    .Replace(
                        $"{canonicalTools} tools with {canonicalOperations} operations",
                        "1 tools with 2 operations",
                        StringComparison.Ordinal));
            var staleValidation = await RunPowerShellScriptAsync(
                Path.Combine(sandbox, "scripts", "check-doc-counts.ps1"),
                ["-SkipBuild"],
                sandbox);
            Assert.NotEqual(0, staleValidation.ExitCode);
            Assert.Contains(
                $"README.md: tool count is 1 but should be {canonicalTools}",
                staleValidation.CombinedOutput,
                StringComparison.Ordinal);

            var allowStale = await RunPowerShellScriptAsync(
                Path.Combine(sandbox, "scripts", "check-doc-counts.ps1"),
                ["-SkipBuild", "-AllowStaleAdvertisedCounts"],
                sandbox);
            Assert.True(allowStale.ExitCode == 0, allowStale.CombinedOutput);

            var incompatible = await RunPowerShellScriptAsync(
                Path.Combine(sandbox, "scripts", "check-doc-counts.ps1"),
                ["-SkipBuild", "-Update", "-AllowStaleAdvertisedCounts"],
                sandbox);
            Assert.NotEqual(0, incompatible.ExitCode);
            Assert.Contains("cannot be used together", incompatible.CombinedOutput, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(sandbox, recursive: true);
        }
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

    private static void CopyDocumentationCountFiles(
        string sandbox,
        int canonicalTools,
        int canonicalOperations)
    {
        var relativePaths = new[]
        {
            "README.md",
            "FEATURES.md",
            "Directory.Build.props",
            Path.Combine("scripts", "check-doc-counts.ps1"),
            Path.Combine("src", "ExcelMcp.McpServer", "README.md"),
            Path.Combine("src", "ExcelMcp.CLI", "README.md"),
            Path.Combine("vscode-extension", "README.md"),
            Path.Combine("vscode-extension", "package.json"),
            Path.Combine("mcpb", "README.md"),
            Path.Combine("mcpb", "manifest.json"),
            Path.Combine("mcpb", "BUILD.md"),
            Path.Combine("gh-pages", "docs", "index.md"),
            Path.Combine("gh-pages", "docs", "faq.md"),
            Path.Combine("gh-pages", "hooks.py"),
            Path.Combine(".github", "plugins", "excel-mcp", "README.md"),
            Path.Combine(".github", "plugins", "excel-cli", "README.md"),
            Path.Combine("skills", "excel-mcp", "SKILL.md"),
            Path.Combine("docs", "INSTALLATION-CLI.md"),
            Path.Combine("docs", "guides", "EXCEL-COM-VS-FILE-PARSERS.md"),
            Path.Combine("docs", "COPILOT-PLUGIN-DISTRIBUTION.md"),
            Path.Combine("src", "ExcelMcp.McpServer", ".mcp", "server.json")
        };

        foreach (var relativePath in relativePaths)
        {
            CopyFile(RepoRoot, sandbox, relativePath);
        }

        var featureRoot = Path.Combine(RepoRoot, "docs", "features");
        foreach (var sourcePath in Directory.GetFiles(featureRoot, "*.md"))
        {
            CopyFile(RepoRoot, sandbox, Path.GetRelativePath(RepoRoot, sourcePath));
        }

        WriteFile(
            sandbox,
            Path.Combine("src", "ExcelMcp.Core", "Models", "Actions", "ToolActions.cs"),
            """
            enum FileAction {
                [JsonStringEnumMemberName("open")] Open,
                [JsonStringEnumMemberName("close")] Close
            }
            """);
        WriteFile(
            sandbox,
            Path.Combine("src", "ExcelMcp.Core", "obj", "GeneratedFiles", "ExcelMcp.Generators",
                "Sbroenne.ExcelMcp.Generators.ServiceRegistryGenerator", "_SkillManifest.g.cs"),
            $$"""
            public static class SkillManifest {
                public const string Json = @"{""TotalCommands"":{{canonicalTools}},""TotalOperations"":{{canonicalOperations - 1}},""Commands"":[{""Name"":""diag"",""Actions"":[""self-test""]}]}";
            }
            """);
        WriteFile(
            sandbox,
            Path.Combine("src", "ExcelMcp.McpServer", "Tools.cs"),
            string.Join(
                Environment.NewLine,
                Enumerable.Range(1, canonicalTools - 1)
                    .Select(index => $"[McpServerTool(Name = \"tool-{index}\")]")) +
                Environment.NewLine +
                "[McpServerTool(Name = \"file\")]");
        WriteFile(
            sandbox,
            Path.Combine("src", "ExcelMcp.McpServer", "Program.cs"),
            """$"Provides {McpToolSurface.ToolCount} tools with {McpToolSurface.OperationCount} operations";""");
        WriteFile(
            sandbox,
            Path.Combine("src", "ExcelMcp.McpServer", "ExcelMcp.McpServer.csproj"),
            """<GenerateSkillFile ExtraOperationCount="2" />""");
        WriteFile(
            sandbox,
            Path.Combine("src", "ExcelMcp.CLI", "ExcelMcp.CLI.csproj"),
            $"""<GenerateSkillFile ExtraOperationCount="2" Description="{canonicalOperations} operations across Excel automation" />""");
    }

    private static void CopyFile(string sourceRoot, string destinationRoot, string relativePath)
    {
        var destinationPath = Path.Combine(destinationRoot, relativePath);
        Directory.CreateDirectory(Path.GetDirectoryName(destinationPath)!);
        File.Copy(Path.Combine(sourceRoot, relativePath), destinationPath);
    }

    private static void WriteFile(string root, string relativePath, string content)
    {
        var path = Path.Combine(root, relativePath);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, content);
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

    private static Dictionary<string, string> ExtractWorkflowJobs(string workflow)
    {
        var normalized = workflow.Replace("\r\n", "\n", StringComparison.Ordinal);
        var jobsMarker = "\njobs:\n";
        var jobsStart = normalized.IndexOf(jobsMarker, StringComparison.Ordinal);
        Assert.True(jobsStart >= 0, "The workflow does not declare any jobs.");

        var jobNames = System.Text.RegularExpressions.Regex
            .Matches(normalized[jobsStart..], @"(?m)^  (?<name>[A-Za-z0-9_-]+):$")
            .Select(match => match.Groups["name"].Value)
            .ToArray();

        Assert.NotEmpty(jobNames);

        return jobNames.ToDictionary(
            name => name,
            name => ExtractWorkflowJob(normalized, name),
            StringComparer.Ordinal);
    }

    private static string ExtractWorkflowJob(string workflow, string jobName)
    {
        workflow = workflow.Replace("\r\n", "\n", StringComparison.Ordinal);
        var marker = $"\n  {jobName}:\n";
        var start = workflow.IndexOf(marker, StringComparison.Ordinal);
        Assert.True(start >= 0, $"Workflow job '{jobName}' was not found.");

        start += marker.Length;
        var end = workflow.IndexOf("\n  ", start, StringComparison.Ordinal);
        while (end >= 0)
        {
            var nextLineEnd = workflow.IndexOf('\n', end + 1);
            var line = nextLineEnd >= 0 ? workflow[end..nextLineEnd] : workflow[end..];
            if (line.Length > 3 && line[3] != ' ')
            {
                break;
            }

            end = workflow.IndexOf("\n  ", end + 3, StringComparison.Ordinal);
        }

        return end >= 0 ? workflow[start..end] : workflow[start..];
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
        IReadOnlyList<string> arguments,
        string? workingDirectory = null)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = "pwsh",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            WorkingDirectory = workingDirectory ?? RepoRoot
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
