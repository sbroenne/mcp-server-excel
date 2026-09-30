using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

/// <summary>
/// Integration tests for release metadata synchronization and workflow wiring.
/// </summary>
[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
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
    private static readonly string TestMcpRegistryPublicationScript = Path.Combine(
        RepoRoot,
        "scripts",
        "Test-McpRegistryPublication.ps1");
    private static readonly string ResolveMcpRegistryReleaseScript = Path.Combine(
        RepoRoot,
        "scripts",
        "Resolve-McpRegistryRelease.ps1");
    private static readonly string ReleaseWorkflow = Path.Combine(
        RepoRoot,
        ".github",
        "workflows",
        "release.yml");
    private static readonly string McpRegistryWorkflow = Path.Combine(
        RepoRoot,
        ".github",
        "workflows",
        "publish-mcp-registry.yml");

    [Theory]
    [InlineData(true, true, false)]
    [InlineData(true, false, false)]
    [InlineData(true, true, true)]
    [InlineData(false, false, false)]
    [Trait("Feature", "ReleaseMetadata")]
    public async Task McpRegistryValidation_ChecksDeclaredRuntimesAndDecodesNuGetReadme(
        bool sourceHasArm64, bool publishedHasArm64, bool wrongArm64Version)
    {
        var sandbox = CreateSandbox();
        try
        {
            var manifestPath = Path.Combine(sandbox, "server.json");
            File.Copy(
                Path.Combine(RepoRoot, "src", "ExcelMcp.McpServer", ".mcp", "server.json"),
                manifestPath);
            var fixtureVersion = ReadJsonProperty(manifestPath, "version");
            var launcherManifestPath = Path.Combine(sandbox, "package.json");
            var launcherManifest = JsonNode.Parse(File.ReadAllText(
                Path.Combine(RepoRoot, "npm-packages", "mcp-server-excel", "package.json")))!.AsObject();
            if (!sourceHasArm64)
            {
                launcherManifest["optionalDependencies"]!.AsObject().Remove("@sbroenne/mcp-server-excel-win32-arm64");
            }
            File.WriteAllText(launcherManifestPath, launcherManifest.ToJsonString());
            var runner = Path.Combine(sandbox, "run.ps1");
            File.WriteAllText(runner, $$"""
                function Invoke-WebRequest {
                    [pscustomobject]@{
                        Content = [Text.Encoding]::UTF8.GetBytes('<!-- mcp-name: io.github.sbroenne/mcp-server-excel -->')
                    }
                }
                function Invoke-RestMethod {
                    param([string]$Uri)
                    if ($Uri -like '*nuget.org*') {
                        if ($Uri -like '*registration5-semver1*') {
                            return [pscustomobject]@{ catalogEntry = 'https://api.nuget.org/v3/catalog0/fixture.json' }
                        }
                        return [pscustomobject]@{
                            id = 'Sbroenne.ExcelMcp.McpServer'
                            version = '{{fixtureVersion}}'
                        }
                    }
                    if ($Uri -like '*win32-*') {
                        $arch = if ($Uri -like '*win32-arm64*') { 'arm64' } else { 'x64' }
                        $runtimeVersion = if ($arch -eq 'arm64' -and ${{wrongArm64Version.ToString().ToLowerInvariant()}}) { '0.0.0' } else { '{{fixtureVersion}}' }
                        return [pscustomobject]@{ name = "@sbroenne/mcp-server-excel-win32-$arch"; version = $runtimeVersion }
                    }
                    $dependencies = @{ '@sbroenne/mcp-server-excel-win32-x64' = '{{fixtureVersion}}' }
                    if (${{publishedHasArm64.ToString().ToLowerInvariant()}}) {
                        $dependencies['@sbroenne/mcp-server-excel-win32-arm64'] = '{{fixtureVersion}}'
                    }
                    return [pscustomobject]@{
                        name = '@sbroenne/mcp-server-excel'
                        version = '{{fixtureVersion}}'
                        mcpName = 'io.github.sbroenne/mcp-server-excel'
                        optionalDependencies = [pscustomobject]$dependencies
                    }
                }
                & '{{TestMcpRegistryPublicationScript.Replace("'", "''", StringComparison.Ordinal)}}' `
                    -ServerJsonPath '{{manifestPath.Replace("'", "''", StringComparison.Ordinal)}}' `
                    -NpmLauncherManifestPath '{{launcherManifestPath.Replace("'", "''", StringComparison.Ordinal)}}' `
                    -Version '{{fixtureVersion}}' -Attempts 1 -RetrySeconds 0
                """);

            var result = await RunPowerShellScriptAsync(runner, [], sandbox);

            if (sourceHasArm64 && (!publishedHasArm64 || wrongArm64Version))
            {
                Assert.NotEqual(0, result.ExitCode);
                Assert.Contains("Published arm64 npm runtime metadata is not ready.", result.CombinedOutput, StringComparison.Ordinal);
            }
            else
            {
                Assert.True(result.ExitCode == 0, result.CombinedOutput);
                Assert.Contains(
                    $"Validated MCP Registry source, NuGet, and npm metadata for version {fixtureVersion}.",
                    result.Stdout,
                    StringComparison.Ordinal);
            }
        }
        finally { Directory.Delete(sandbox, recursive: true); }
    }

    [Fact]
    [Trait("Feature", "ReleaseMetadata")]
    public void McpRegistryRepair_UsesOnlyExactExistingRelease()
    {
        var release = File.ReadAllText(ReleaseWorkflow);
        var registry = File.ReadAllText(McpRegistryWorkflow);
        var publishMcpRegistry = ExtractWorkflowJob(release, "publish-mcp-registry");

        Assert.Contains("uses: ./.github/workflows/publish-mcp-registry.yml", publishMcpRegistry, StringComparison.Ordinal);
        Assert.Contains("needs: [version, create-tag, create-release, publish]", publishMcpRegistry, StringComparison.Ordinal);
        Assert.Contains("workflow_dispatch:", registry, StringComparison.Ordinal);
        Assert.Contains("environment: mcp-registry", registry, StringComparison.Ordinal);
        Assert.Contains("./scripts/Resolve-McpRegistryRelease.ps1", registry, StringComparison.Ordinal);
        Assert.Contains("ref: ${{ github.sha }}", registry, StringComparison.Ordinal);
        Assert.Contains("ref: ${{ needs.resolve.outputs.commit }}", registry, StringComparison.Ordinal);
        Assert.Contains("./automation/scripts/Test-McpRegistryPublication.ps1", registry, StringComparison.Ordinal);
        Assert.Contains("source/src/ExcelMcp.McpServer/.mcp/server.json", registry, StringComparison.Ordinal);
        Assert.DoesNotContain("dotnet build", registry, StringComparison.Ordinal);
        Assert.DoesNotContain("npm publish", registry, StringComparison.Ordinal);
        Assert.DoesNotContain("gh release create", registry, StringComparison.Ordinal);
        Assert.DoesNotContain("gh release upload", registry, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    [Trait("Feature", "ReleaseMetadata")]
    public async Task McpRegistryRepair_RequiresTagCommitOnMain(bool tagCommitOnMain)
    {
        var sandbox = CreateSandbox();
        try
        {
            var repository = Path.Combine(sandbox, "source");
            var remote = Path.Combine(sandbox, "remote.git");
            Directory.CreateDirectory(repository);
            var runner = Path.Combine(sandbox, "run.ps1");
            File.WriteAllText(runner, $$"""
                $ErrorActionPreference = 'Stop'
                git init --bare '{{remote.Replace("'", "''", StringComparison.Ordinal)}}'
                git init '{{repository.Replace("'", "''", StringComparison.Ordinal)}}'
                Set-Location '{{repository.Replace("'", "''", StringComparison.Ordinal)}}'
                git config user.name fixture
                git config user.email fixture@example.test
                Set-Content release.txt base
                git add release.txt
                git commit -m base
                git branch -M main
                git remote add origin '{{remote.Replace("'", "''", StringComparison.Ordinal)}}'
                git push -u origin main
                if (-not ${{tagCommitOnMain.ToString().ToLowerInvariant()}}) {
                    git checkout -b unmerged
                    Set-Content release.txt unmerged
                    git commit -am unmerged
                }
                git tag v1.2.3
                $env:GITHUB_OUTPUT = Join-Path $PWD output.txt
                function gh {
                    [pscustomobject]@{ isDraft = $false; tagName = 'v1.2.3' } | ConvertTo-Json -Compress
                }
                & '{{ResolveMcpRegistryReleaseScript.Replace("'", "''", StringComparison.Ordinal)}}' -Tag v1.2.3
                """);

            var result = await RunPowerShellScriptAsync(runner, [], sandbox);

            Assert.Equal(tagCommitOnMain, result.ExitCode == 0);
            if (tagCommitOnMain)
            {
                Assert.Contains("version=1.2.3", File.ReadAllText(Path.Combine(repository, "output.txt")), StringComparison.Ordinal);
            }
            else
            {
                Assert.Contains("not reachable from protected main", result.Stderr, StringComparison.Ordinal);
                Assert.False(File.Exists(Path.Combine(repository, "output.txt")));
            }
        }
        finally { DeleteGitSandbox(sandbox); }
    }

    [Theory]
    [InlineData("vscode-extension", "package-lock.json")]
    [InlineData("src/ExcelMcp.McpServer/.mcp", "server.json")]
    [Trait("Feature", "ReleaseMetadata")]
    public async Task UpdateReleaseVersion_MissingMetadataCannotPartiallyStampOtherFiles(string directory, string file)
    {
        var sandbox = CreateSandbox();
        try
        {
            CopyReleaseMetadataFiles(sandbox);
            var before = File.ReadAllText(Path.Combine(sandbox, "package.json"));
            File.Delete(Path.Combine(sandbox, directory, file));
            var result = await RunPowerShellScriptAsync(UpdateReleaseVersionScript,
                ["-RepoRoot", sandbox, "-Version", "9.8.7"]);
            Assert.NotEqual(0, result.ExitCode);
            Assert.Equal(before, File.ReadAllText(Path.Combine(sandbox, "package.json")));
        }
        finally { Directory.Delete(sandbox, recursive: true); }
    }

    [Fact]
    [Trait("Feature", "ReleaseMetadata")]
    public void PluginPublication_UsesExplicitReleaseIdentityAndPreparedInputs()
    {
        var release = File.ReadAllText(ReleaseWorkflow);
        var plugins = File.ReadAllText(Path.Combine(RepoRoot, ".github", "workflows", "publish-plugins.yml"));
        Assert.DoesNotContain("workflow_run", plugins, StringComparison.Ordinal);
        Assert.Contains("workflow_call:", plugins, StringComparison.Ordinal);
        Assert.Contains("needs: [version, create-tag, create-release, publish]", ExtractWorkflowJob(release, "publish-plugins"), StringComparison.Ordinal);
        Assert.Contains("release_commit:", plugins, StringComparison.Ordinal);
        Assert.Contains("plugin_artifact:", plugins, StringComparison.Ordinal);
        Assert.Contains("ref: ${{ needs.resolve.outputs.commit }}", plugins, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("1.2.2", false, false, "1.2.3", false, true)]
    [InlineData("1.2.3", true, false, "1.2.3", false, true)]
    [InlineData("1.2.3", true, true, "1.2.3", false, true)]
    [InlineData("1.2.4", false, false, "1.2.3", false, false)]
    [InlineData("1.2.2", true, false, "1.2.3", false, false)]
    [InlineData("1.2.2", false, false, "9.9.9", false, false)]
    [InlineData("1.2.2", false, false, "1.2.3", true, false)]
    [Trait("Feature", "ReleaseMetadata")]
    public async Task PluginPublication_ExercisesVersionRetryAndFailureDecisions(
        string publishedVersion, bool tagExists, bool manualRepair, string payloadVersion, bool syncFails, bool succeeds)
    {
        var sandbox = CreateSandbox();
        try
        {
            var manifest = new { plugins = new[] { new { version = publishedVersion } } };
            WriteFile(sandbox, Path.Combine("published-repo", "marketplace.json"), JsonSerializer.Serialize(manifest));
            foreach (var name in new[] { "excel-cli", "excel-mcp" })
            {
                WriteFile(sandbox, Path.Combine("built-plugins", name, "plugin.json"),
                    JsonSerializer.Serialize(new { name, version = payloadVersion }));
                WriteFile(sandbox, Path.Combine("built-plugins", name, "version.txt"), "1.2.3");
                WriteFile(sandbox, Path.Combine("built-plugins", name, "skills", name, "VERSION"), "1.2.3");
            }
            WriteFile(sandbox, Path.Combine("source", "scripts", "Sync-PublishedPluginRepo.ps1"),
                syncFails ? "throw 'sync-root-cause'" : "'synced' | Set-Content sync.txt");
            var workflow = File.ReadAllText(Path.Combine(RepoRoot, ".github", "workflows", "publish-plugins.yml"));
            var step = ExtractPowerShellStep(workflow, "Guard, synchronize and publish");
            var runner = Path.Combine(sandbox, "run.ps1");
            File.WriteAllText(runner, $$"""
                $env:VERSION='1.2.3'
                $env:TAG='v1.2.3'
                $env:SOURCE_COMMIT='exact-released-commit'
                $env:MANUAL_REPAIR='{{manualRepair.ToString().ToLowerInvariant()}}'
                $env:GITHUB_STEP_SUMMARY=Join-Path $PWD summary.txt
                function git {
                    $global:LASTEXITCODE=0
                    if ($args -contains '--list') {
                        if ({{(tagExists ? "$true" : "$false")}}) { 'v1.2.3' }
                        return
                    }
                    if ($args -contains 'diff') { 'plugins/example'; return }
                    Add-Content git-calls.txt ($args -join ' ')
                }
                {{step}}
                """);
            var result = await RunPowerShellScriptAsync(runner, [], sandbox);
            Assert.True((result.ExitCode == 0) == succeeds, result.CombinedOutput);
            if (!succeeds)
            {
                var expectedError = publishedVersion == "1.2.4" ? "Downgrade publish blocked"
                    : tagExists ? "Existing tag conflicts"
                    : payloadVersion != "1.2.3" ? "Wrong identity"
                    : "sync-root-cause";
                Assert.Contains(expectedError, result.Stderr, StringComparison.Ordinal);
            }
            var skipped = tagExists && !manualRepair && succeeds;
            Assert.Equal(succeeds && !skipped, File.Exists(Path.Combine(sandbox, "sync.txt")));
            var callsFile = Path.Combine(sandbox, "git-calls.txt");
            if (succeeds && !skipped)
            {
                var calls = File.ReadAllText(callsFile);
                Assert.Contains("Source release commit: exact-released-commit", calls, StringComparison.Ordinal);
                Assert.Contains("push origin HEAD:main", calls, StringComparison.Ordinal);
                Assert.Equal(!tagExists, calls.Contains("push origin v1.2.3", StringComparison.Ordinal));
                Assert.DoesNotContain("--force", calls, StringComparison.Ordinal);
            }
            else { Assert.False(File.Exists(callsFile)); }
        }
        finally { Directory.Delete(sandbox, recursive: true); }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    [Trait("Feature", "ReleaseMetadata")]
    public async Task PluginRepair_VerifiesPayloadChecksumBeforeExtracting(bool validChecksum)
    {
        var sandbox = CreateSandbox();
        try
        {
            WriteFile(sandbox, Path.Combine("payload", "release.txt"), "exact released payload");
            Directory.CreateDirectory(Path.Combine(sandbox, "downloads"));
            var zip = Path.Combine(sandbox, "downloads", "excel-plugins-v1.2.3.zip");
            System.IO.Compression.ZipFile.CreateFromDirectory(Path.Combine(sandbox, "payload"), zip);
            var hash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(zip)));
            WriteFile(sandbox, Path.Combine("downloads", "SHA256SUMS"),
                $"{(validChecksum ? hash : new string('0', 64))}  excel-plugins-v1.2.3.zip\n");
            var workflow = File.ReadAllText(Path.Combine(RepoRoot, ".github", "workflows", "publish-plugins.yml"));
            var step = ExtractPowerShellStep(workflow, "Recover or rebuild only the requested release");
            var runner = Path.Combine(sandbox, "run.ps1");
            File.WriteAllText(runner, $$"""
                $env:VERSION='1.2.3'
                $env:TAG='v1.2.3'
                $env:GITHUB_REPOSITORY='fixture/source'
                function gh {
                    $global:LASTEXITCODE=0
                    if ($args -notcontains 'v1.2.3') { throw 'Wrong release selected.' }
                    if ($args -contains 'view') { '{"assets":[{"name":"excel-plugins-v1.2.3.zip"}]}'; return }
                    if ($args -contains 'download') { Copy-Item downloads\* . -Force; return }
                    throw "Unexpected remote command: $args"
                }
                {{step}}
                """);
            var result = await RunPowerShellScriptAsync(runner, [], sandbox);
            Assert.True((result.ExitCode == 0) == validChecksum, result.CombinedOutput);
            Assert.Equal(validChecksum, File.Exists(Path.Combine(sandbox, "built-plugins", "release.txt")));
            if (!validChecksum) { Assert.Contains("checksum", result.Stderr, StringComparison.OrdinalIgnoreCase); }
        }
        finally { Directory.Delete(sandbox, recursive: true); }
    }

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
        Assert.Contains("./scripts/Build-Changelog.ps1", releaseWorkflow, StringComparison.Ordinal);
        Assert.Contains("git apply prepared-release/release-metadata.patch", releaseWorkflow, StringComparison.Ordinal);
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
        var buildPackages = ExtractWorkflowJob(releaseWorkflow, "build-packages");
        var createTag = ExtractWorkflowJob(releaseWorkflow, "create-tag");
        var createRelease = ExtractWorkflowJob(releaseWorkflow, "create-release");

        Assert.Contains("./scripts/Build-Changelog.ps1", prepareRelease, StringComparison.Ordinal);
        Assert.Contains("name: release-metadata", prepareRelease, StringComparison.Ordinal);
        Assert.Contains("needs: [version, prepare-release]", buildPackages, StringComparison.Ordinal);
        Assert.Contains("name: release-metadata", buildPackages, StringComparison.Ordinal);
        Assert.Contains("git apply prepared-release/release-metadata.patch", buildPackages, StringComparison.Ordinal);
        Assert.Contains("git apply prepared-release/release-metadata.patch", createTag, StringComparison.Ordinal);
        Assert.Contains("Build-ReleasePackages.ps1", buildPackages, StringComparison.Ordinal);
        Assert.DoesNotContain("dotnet publish", releaseWorkflow, StringComparison.Ordinal);

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
        var registryValidation = File.ReadAllText(TestMcpRegistryPublicationScript);
        Assert.Contains("@sbroenne%2fmcp-server-excel-win32-$architecture/$Version", registryValidation, StringComparison.Ordinal);
        Assert.Contains("$sourceLauncher.optionalDependencies", registryValidation, StringComparison.Ordinal);
        Assert.Contains("$launcher.optionalDependencies.$runtimeName -ne $Version", registryValidation, StringComparison.Ordinal);
        Assert.Contains("$launcher.mcpName", registryValidation, StringComparison.Ordinal);
        Assert.Contains("$runtime.version -ne $Version", registryValidation, StringComparison.Ordinal);
        Assert.Contains("$readme -notmatch", registryValidation, StringComparison.Ordinal);
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "ReleaseMetadata")]
    public void ReleaseFlow_BuildsTestsAndPublishesBothNpmDistributions()
    {
        var workflow = File.ReadAllText(ReleaseWorkflow);
        var build = ExtractWorkflowJob(workflow, "build-packages");
        var publish = ExtractWorkflowJob(workflow, "publish");

        Assert.Contains("actions/setup-node@", build, StringComparison.Ordinal);
        Assert.Contains("Build-ReleasePackages.ps1", build, StringComparison.Ordinal);
        Assert.Contains("name: release-packages", publish, StringComparison.Ordinal);
        Assert.Contains("Detect npm authentication mode", publish, StringComparison.Ordinal);
        Assert.Contains("if: steps.npm-auth.outputs.mode == 'token'", publish, StringComparison.Ordinal);
        Assert.Contains("if: steps.npm-auth.outputs.mode == 'oidc'", publish, StringComparison.Ordinal);
        Assert.Contains("NPM_BOOTSTRAP_TOKEN: ${{ secrets.NPM_TOKEN }}", publish, StringComparison.Ordinal);
        Assert.Contains("$env:NODE_AUTH_TOKEN = $env:NPM_BOOTSTRAP_TOKEN", publish, StringComparison.Ordinal);
        Assert.DoesNotContain("NODE_AUTH_TOKEN: ${{ secrets.NPM_TOKEN }}", publish, StringComparison.Ordinal);
        Assert.Contains("npm publish \"./artifacts/npm/sbroenne-$name-$env:VERSION.tgz\" --access public", publish, StringComparison.Ordinal);

        foreach (var packageName in new[] { "excelcli", "mcp-server-excel" })
        {
            var launcherIndex = publish.IndexOf(
                $"'{packageName}'",
                StringComparison.Ordinal);
            foreach (var architecture in new[] { "x64", "arm64" })
            {
                var runtimeIndex = publish.IndexOf(
                    $"'{packageName}-win32-{architecture}'",
                    StringComparison.Ordinal);
                Assert.True(runtimeIndex >= 0);
                Assert.True(launcherIndex > runtimeIndex, "Publish both runtimes before their launcher.");
            }
        }

        var preCommit = File.ReadAllText(Path.Combine(RepoRoot, "scripts", "pre-commit.ps1"));
        Assert.DoesNotContain("Build-NpmPackages.ps1", preCommit, StringComparison.Ordinal);
        Assert.DoesNotContain("Test-NpmPackages.ps1", preCommit, StringComparison.Ordinal);
        var packages = File.ReadAllText(Path.Combine(RepoRoot, "scripts", "Build-ReleasePackages.ps1"));
        Assert.Contains("Build-NpmPackages.ps1", packages, StringComparison.Ordinal);
        Assert.Contains("Test-NpmPackages.ps1", packages, StringComparison.Ordinal);
        var ci = File.ReadAllText(Path.Combine(RepoRoot, ".github", "workflows", "ci.yml"));
        Assert.Contains("Build-ReleasePackages.ps1", ci, StringComparison.Ordinal);
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "ReleaseMetadata")]
    public void ReleaseFlow_DoesNotGenerateOrPatchDocumentationCounts()
    {
        var releaseWorkflow = File.ReadAllText(ReleaseWorkflow);

        Assert.DoesNotContain("check-doc-counts.ps1", releaseWorkflow, StringComparison.Ordinal);
        Assert.DoesNotContain("release-doc-counts.patch", releaseWorkflow, StringComparison.Ordinal);
        Assert.DoesNotContain("Apply Generated Documentation Counts", releaseWorkflow, StringComparison.Ordinal);

        Assert.False(File.Exists(Path.Combine(RepoRoot, ".github", "workflows", "doc-counts.yml")));
        var ci = File.ReadAllText(Path.Combine(RepoRoot, ".github", "workflows", "ci.yml"));
        Assert.Contains("check-doc-counts.ps1 -SkipBuild", ci, StringComparison.Ordinal);
        Assert.DoesNotContain("-AllowStaleAdvertisedCounts", ci, StringComparison.Ordinal);
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "ReleaseMetadata")]
    public void DocumentationCounts_DefaultRefreshBuildsProjectDependencies()
    {
        // A fresh checkout has no built Service/ComInterop outputs, so the MCP Server
        // build must include its project dependencies.
        var script = File.ReadAllText(Path.Combine(RepoRoot, "scripts", "check-doc-counts.ps1"));

        Assert.DoesNotContain("--no-dependencies", script, StringComparison.Ordinal);
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
            var hooksContent = await File.ReadAllTextAsync(hooksPath);
            Assert.Contains("_read_release_headline_counts()", hooksContent, StringComparison.Ordinal);
            Assert.Contains("for output_name, source_rel in FEATURE_SOURCES.items():", hooksContent, StringComparison.Ordinal);
            Assert.DoesNotMatch(
                @"exposing \d+ tools and \d+ operations",
                hooksContent);

            var docCountsPath = Path.Combine(sandbox, "doc-counts.json");
            Assert.True(File.Exists(docCountsPath), "-Update must generate the single doc-counts.json include file.");
            Assert.Equal(canonicalTools, ReadJsonInt(docCountsPath, "tools"));
            Assert.Equal(canonicalOperations, ReadJsonInt(docCountsPath, "operations"));

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

            // Restore the README headline, then make doc-counts.json itself stale to
            // prove it is validated (and regenerated) independently of the headline text.
            await File.WriteAllTextAsync(
                readmePath,
                (await File.ReadAllTextAsync(readmePath))
                    .Replace(
                        "1 tools with 2 operations",
                        $"{canonicalTools} tools with {canonicalOperations} operations",
                        StringComparison.Ordinal));
            await File.WriteAllTextAsync(docCountsPath, "{\n  \"tools\": 1,\n  \"operations\": 2\n}\n");

            var staleDocCounts = await RunPowerShellScriptAsync(
                Path.Combine(sandbox, "scripts", "check-doc-counts.ps1"),
                ["-SkipBuild"],
                sandbox);
            Assert.NotEqual(0, staleDocCounts.ExitCode);
            Assert.Contains("doc-counts.json", staleDocCounts.CombinedOutput, StringComparison.Ordinal);

            var allowStaleDocCounts = await RunPowerShellScriptAsync(
                Path.Combine(sandbox, "scripts", "check-doc-counts.ps1"),
                ["-SkipBuild", "-AllowStaleAdvertisedCounts"],
                sandbox);
            Assert.True(allowStaleDocCounts.ExitCode == 0, allowStaleDocCounts.CombinedOutput);

            var regenerate = await RunPowerShellScriptAsync(
                Path.Combine(sandbox, "scripts", "check-doc-counts.ps1"),
                ["-Update", "-SkipBuild"],
                sandbox);
            Assert.True(regenerate.ExitCode == 0, regenerate.CombinedOutput);
            Assert.Equal(canonicalTools, ReadJsonInt(docCountsPath, "tools"));
            Assert.Equal(canonicalOperations, ReadJsonInt(docCountsPath, "operations"));

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
            var nugetPackage = Assert.Single(packages, package =>
                package.GetProperty("identifier").GetString() == "Sbroenne.ExcelMcp.McpServer");
            Assert.Equal("9.8.7", nugetPackage.GetProperty("version").GetString());

            var npmPackage = Assert.Single(packages, package =>
                package.GetProperty("identifier").GetString() == "@sbroenne/mcp-server-excel");
            Assert.Equal("9.8.7", npmPackage.GetProperty("version").GetString());
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

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "ReleaseMetadata")]
    public async Task UpdateMetadata_RejectsMissingNpmPackage()
    {
        var sandbox = CreateSandbox();
        try
        {
            var metadataPath = Path.Combine(sandbox, "server.json");
            await File.WriteAllTextAsync(
                metadataPath,
                """
                {
                  "version": "1.0.0",
                  "packages": [
                    {
                      "identifier": "Sbroenne.ExcelMcp.McpServer",
                      "version": "1.0.0"
                    }
                  ]
                }
                """);

            var result = await RunPowerShellScriptAsync(
                UpdateMetadataScript,
                ["-ServerJsonPath", metadataPath, "-Version", "9.8.7"]);

            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains(
                "@sbroenne/mcp-server-excel",
                result.CombinedOutput,
                StringComparison.Ordinal);
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
            Path.Combine("artifacts", "generated-skills", "excel-mcp", "SKILL.md"),
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

    private static void DeleteGitSandbox(string path)
    {
        foreach (var entry in Directory.EnumerateFileSystemEntries(path, "*", SearchOption.AllDirectories)
                     .OrderByDescending(value => value.Length))
        {
            File.SetAttributes(entry, FileAttributes.Normal);
        }
        File.SetAttributes(path, FileAttributes.Normal);
        Directory.Delete(path, recursive: true);
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
            ReadJsonProperty(
                Path.Combine(root, "src", "ExcelMcp.McpServer", ".mcp", "server.json"),
                "packages",
                "1",
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

    private static int ReadJsonInt(string path, string propertyName)
    {
        using var document = JsonDocument.Parse(File.ReadAllText(path));
        return document.RootElement.GetProperty(propertyName).GetInt32();
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

    private static string ExtractPowerShellStep(string workflow, string stepName)
    {
        var lines = workflow.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n');
        var start = Array.FindIndex(lines, line => line == $"      - name: {stepName}");
        Assert.True(start >= 0, $"Step '{stepName}' was not found.");
        var run = Array.FindIndex(lines, start, line => line == "        run: |");
        Assert.True(run > start);
        var script = lines.Skip(run + 1).TakeWhile(line => line.Length == 0 || line.StartsWith("          ", StringComparison.Ordinal))
            .Select(line => line.Length == 0 ? "" : line[10..]);
        return string.Join(Environment.NewLine, script);
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
        try { await process.WaitForExitAsync(timeout.Token); }
        catch (OperationCanceledException)
        {
            process.Kill(entireProcessTree: true);
            await process.WaitForExitAsync();
            throw new TimeoutException("Release metadata regression exceeded 30 seconds.");
        }

        return new ScriptResult(process.ExitCode, await stdout, await stderr);
    }

    private sealed record ScriptResult(int ExitCode, string Stdout, string Stderr)
    {
        public string CombinedOutput => $"{Stdout}{Environment.NewLine}{Stderr}";
    }
}
