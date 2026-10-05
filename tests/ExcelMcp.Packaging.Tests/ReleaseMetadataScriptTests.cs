using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using Xunit;

namespace Sbroenne.ExcelMcp.Packaging.Tests;

/// <summary>
/// Integration tests for release metadata synchronization and workflow wiring.
/// </summary>
[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
public sealed class ReleaseMetadataScriptTests
{
    private static readonly string RepoRoot = FindRepoRoot();
    private static readonly string[] GitRepositoryEnvironmentVariables = GetGitRepositoryEnvironmentVariables();
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
    [InlineData("absent", false, true)]
    [InlineData("draft", false, true)]
    [InlineData("draft-later-page", false, true)]
    [InlineData("duplicate-draft", false, false)]
    [InlineData("invalid-draft", false, false)]
    [InlineData("published-list", false, false)]
    [InlineData("list-api-error", false, false)]
    [InlineData("published", false, true)]
    [InlineData("immutable", false, true)]
    [InlineData("missing", false, false)]
    [InlineData("missing", true, true)]
    [InlineData("immutable-missing", true, false)]
    [InlineData("mismatch", true, false)]
    [InlineData("api-error", false, false)]
    [Trait("Feature", "ReleaseMetadata")]
    public async Task GitHubRelease_VerifiesPublishedAssetsAndOnlyReplacesDraftAssets(
        string mode, bool allowMutableRepair, bool succeeds)
    {
        var sandbox = CreateSandbox();
        try
        {
            var artifacts = Path.Combine(sandbox, "artifacts");
            Directory.CreateDirectory(artifacts);
            foreach (var name in new[]
            {
                "ExcelMcp-CLI-1.2.3-windows.zip", "ExcelMcp-MCP-Server-1.2.3-windows.zip",
                "excel-plugins-v1.2.3.zip", "excel-skills-v1.2.3.zip",
                "excel-mcp-1.2.3.vsix", "excel-mcp-1.2.3-win32-arm64.vsix", "excel-mcp-1.2.3.mcpb"
            })
            {
                File.WriteAllText(Path.Combine(artifacts, name), name);
            }
            File.WriteAllText(Path.Combine(sandbox, "metadata.patch"), "exact metadata patch");
            var script = Path.Combine(RepoRoot, "scripts", "Publish-GitHubRelease.ps1");
            var runner = Path.Combine(sandbox, "run.ps1");
            File.WriteAllText(runner, $$"""
                $ErrorActionPreference = 'Stop'
                $global:releaseFixtureCreated = $false
                $global:releaseFixtureUploaded = $false
                $global:releaseFixturePublished = $false
                $global:releaseFixtureCalls = [Collections.Generic.List[string]]::new()
                function global:gh {
                    param([Parameter(ValueFromRemainingArguments)][string[]]$Arguments)
                    $global:releaseFixtureCalls.Add(($Arguments -join ' '))
                    $global:LASTEXITCODE = 0
                    if ($Arguments[0] -eq 'api') {
                        if ('{{mode}}' -eq 'api-error') {
                            $global:LASTEXITCODE = 1
                            'gh: Forbidden (HTTP 403)'
                            return
                        }
                        $isList = $Arguments[1] -eq 'repos/owner/repo/releases'
                        $isAbsent = '{{mode}}' -eq 'absent' -and -not $global:releaseFixtureCreated
                        $isDraft = '{{mode}}' -in @('absent', 'draft', 'draft-later-page',
                            'duplicate-draft', 'invalid-draft', 'published-list', 'list-api-error') `
                            -and -not $global:releaseFixturePublished
                        if (-not $isList -and ($isAbsent -or $isDraft)) {
                            $global:LASTEXITCODE = 1
                            'gh: Not Found (HTTP 404)'
                            return
                        }
                        if ($isList) {
                            if ('--paginate' -notin $Arguments -or '--slurp' -notin $Arguments) {
                                throw 'Draft lookup must request all release pages.'
                            }
                            if ('{{mode}}' -eq 'list-api-error') {
                                $global:LASTEXITCODE = 1
                                'gh: Forbidden (HTTP 403)'
                                return
                            }
                            if ($isAbsent) { '[[]]'; return }
                        }
                        $assets = @(Get-ChildItem publish -File | ForEach-Object {
                            @{ name = $_.Name; digest = 'sha256:' + (Get-FileHash $_.FullName).Hash.ToLowerInvariant() }
                        })
                        if ('{{mode}}' -match 'missing' -and -not $global:releaseFixtureUploaded) {
                            $assets = @($assets | Where-Object name -ne 'SHA256SUMS')
                        }
                        if ('{{mode}}' -eq 'mismatch') { $assets[0].digest = 'sha256:' + ('0' * 64) }
                        $release = @{
                            tag_name = 'v1.2.3'
                            draft = $isDraft
                            immutable = '{{mode}}' -like 'immutable*'
                            assets = $assets
                        }
                        if ($isList) {
                            $pages = ,@($release)
                            if ('{{mode}}' -eq 'draft-later-page') {
                                $pages = @(@(@{ tag_name = 'V1.2.3'; draft = $true; immutable = $false }),
                                    @($release))
                            }
                            if ('{{mode}}' -eq 'duplicate-draft') { $pages = @(@($release), @($release)) }
                            if ('{{mode}}' -eq 'invalid-draft') { $release.immutable = 'false' }
                            if ('{{mode}}' -eq 'published-list') { $release.draft = $false }
                            $response = ConvertTo-Json -InputObject $pages -Depth 7
                            $response | Set-Content list-response.json
                            $response
                        } else {
                            $release | ConvertTo-Json -Depth 5
                        }
                    } elseif ($Arguments[1] -eq 'create') {
                        $global:releaseFixtureCreated = $true
                    } elseif ($Arguments[1] -eq 'upload') {
                        $global:releaseFixtureUploaded = $true
                    } elseif ($Arguments[1] -eq 'edit') {
                        $global:releaseFixturePublished = $true
                    }
                }
                try {
                    & '{{script.Replace("'", "''", StringComparison.Ordinal)}}' -Version 1.2.3 -Repository owner/repo `
                      -SourceCommit ('a' * 40) -ReleaseCommit ('b' * 40) `
                      -MetadataPatch metadata.patch -AssetDirectory artifacts -PublishDirectory publish `
                      -NotesFile metadata.patch -AllowMutableRepair:{{(allowMutableRepair ? "$true" : "$false")}}
                } finally {
                    ConvertTo-Json -InputObject @($global:releaseFixtureCalls.ToArray()) | Set-Content calls.json
                }
                """);
            var result = await RunPowerShellScriptAsync(runner, [], sandbox);
            Assert.True(succeeds == (result.ExitCode == 0), result.CombinedOutput);
            if (mode is "absent" or "draft" or "draft-later-page" or "duplicate-draft" or "invalid-draft" or "published-list")
            {
                using var response = JsonDocument.Parse(File.ReadAllText(Path.Combine(sandbox, "list-response.json")));
                var pages = response.RootElement;
                Assert.Equal(JsonValueKind.Array, pages.ValueKind);
                Assert.All(pages.EnumerateArray(), page =>
                {
                    Assert.Equal(JsonValueKind.Array, page.ValueKind);
                    Assert.Equal(JsonValueKind.Object, Assert.Single(page.EnumerateArray()).ValueKind);
                });
                Assert.Equal(mode is "draft-later-page" or "duplicate-draft" ? 2 : 1, pages.GetArrayLength());
                if (mode == "draft-later-page")
                {
                    Assert.Equal("V1.2.3", pages[0][0].GetProperty("tag_name").GetString());
                    Assert.Equal("v1.2.3", pages[1][0].GetProperty("tag_name").GetString());
                }
            }
            if (!succeeds)
            {
                var expectedError = mode switch
                {
                    "missing" or "immutable-missing" => "Published assets are missing",
                    "mismatch" => "Missing or mismatched GitHub SHA-256 digest",
                    "api-error" or "list-api-error" => "GitHub command failed",
                    "duplicate-draft" => "multiple releases",
                    "invalid-draft" or "published-list" => "invalid release identity or state",
                    _ => throw new InvalidOperationException($"Unexpected failure fixture: {mode}")
                };
                Assert.Contains(expectedError, result.CombinedOutput, StringComparison.Ordinal);
            }
            var calls = File.ReadAllText(Path.Combine(sandbox, "calls.json"));
            if (mode is "absent" or "draft" or "draft-later-page")
            {
                Assert.Contains("release upload", calls, StringComparison.Ordinal);
                Assert.Contains("--clobber", calls, StringComparison.Ordinal);
                Assert.Contains("release edit", calls, StringComparison.Ordinal);
                Assert.Contains("api repos/owner/repo/releases --paginate --slurp", calls, StringComparison.Ordinal);
                Assert.Contains("Published v1.2.3 after verifying every draft asset", result.CombinedOutput, StringComparison.Ordinal);
                if (mode == "absent")
                {
                    Assert.Contains("--draft", calls, StringComparison.Ordinal);
                    Assert.Contains("--verify-tag", calls, StringComparison.Ordinal);
                }
                else
                {
                    Assert.DoesNotContain("release create", calls, StringComparison.Ordinal);
                }
            }
            else
            {
                Assert.DoesNotContain("--clobber", calls, StringComparison.Ordinal);
                Assert.DoesNotContain("release edit", calls, StringComparison.Ordinal);
                Assert.DoesNotContain("release create", calls, StringComparison.Ordinal);
                if (mode == "missing" && succeeds)
                {
                    Assert.Contains("release upload", calls, StringComparison.Ordinal);
                }
                else
                {
                    Assert.DoesNotContain("release upload", calls, StringComparison.Ordinal);
                }
            }
            if (mode is "published" or "immutable" or "missing" or "immutable-missing" or "mismatch" or "api-error")
            {
                Assert.DoesNotContain("api repos/owner/repo/releases --paginate", calls, StringComparison.Ordinal);
            }
            if (succeeds)
            {
                using var inputs = JsonDocument.Parse(File.ReadAllText(Path.Combine(sandbox, "publish", "RELEASE-INPUTS.json")));
                Assert.Equal(".github/workflows/release.yml", inputs.RootElement.GetProperty("workflowPath").GetString());
                Assert.Equal(new string('a', 40), inputs.RootElement.GetProperty("sourceCommit").GetString());
                Assert.Equal(new string('b', 40), inputs.RootElement.GetProperty("releaseCommit").GetString());
                Assert.Equal("v1.2.3", inputs.RootElement.GetProperty("tag").GetString());
                Assert.Equal("build-input-record-not-attestation", inputs.RootElement.GetProperty("kind").GetString());
                Assert.Equal(
                    Convert.ToHexStringLower(SHA256.HashData(File.ReadAllBytes(Path.Combine(sandbox, "metadata.patch")))),
                    inputs.RootElement.GetProperty("metadataPatchSha256").GetString());
                Assert.Equal(7, inputs.RootElement.GetProperty("artifacts").GetArrayLength());
                foreach (var artifact in inputs.RootElement.GetProperty("artifacts").EnumerateArray())
                {
                    var name = artifact.GetProperty("name").GetString()!;
                    Assert.Equal(
                        Convert.ToHexStringLower(SHA256.HashData(File.ReadAllBytes(Path.Combine(artifacts, name)))),
                        artifact.GetProperty("sha256").GetString());
                }
                var checksumLines = File.ReadAllLines(Path.Combine(sandbox, "publish", "SHA256SUMS"));
                Assert.Equal(9, checksumLines.Length);
                foreach (var line in checksumLines)
                {
                    var parts = line.Split("  ", StringSplitOptions.None);
                    Assert.Equal(2, parts.Length);
                    Assert.Equal(
                        Convert.ToHexStringLower(SHA256.HashData(File.ReadAllBytes(Path.Combine(sandbox, "publish", parts[1])))),
                        parts[0]);
                }
            }
        }
        finally
        {
            Directory.Delete(sandbox, recursive: true);
        }
    }

    [Fact]
    [Trait("Feature", "ReleaseMetadata")]
    public void VscodePublication_UsesWindowsForWindowsOnlyExtensionTools()
    {
        var publish = ExtractWorkflowJob(File.ReadAllText(ReleaseWorkflow), "publish-vscode");

        Assert.Contains("needs: [version, create-tag]", publish, StringComparison.Ordinal);
        Assert.DoesNotContain("if:", publish, StringComparison.Ordinal);
        Assert.Contains("ref: ${{ needs.create-tag.outputs.commit }}", publish, StringComparison.Ordinal);
        Assert.Contains("name: release-packages", publish, StringComparison.Ordinal);
        Assert.Contains("runs-on: windows-latest", publish, StringComparison.Ordinal);
        Assert.Contains("npm ci --ignore-scripts", publish, StringComparison.Ordinal);
        Assert.Contains("shell: pwsh", publish, StringComparison.Ordinal);
        Assert.Contains("Publish Windows x64 extension", publish, StringComparison.Ordinal);
        Assert.Contains("Publish Windows ARM64 extension", publish, StringComparison.Ordinal);
        Assert.Contains("--skip-duplicate", publish, StringComparison.Ordinal);
    }

    [Fact]
    [Trait("Feature", "ReleaseMetadata")]
    public void PublicationJobs_UseImmutableActionReferences()
    {
        var workflow = File.ReadAllText(ReleaseWorkflow);
        foreach (var job in new[] { "publish-vscode", "verify-arm64" })
        {
            var body = ExtractWorkflowJob(workflow, job);
            foreach (var line in body.Split('\n').Where(line => line.Contains("uses: actions/", StringComparison.Ordinal)))
            {
                Assert.Matches(@"uses: actions/[a-z-]+@[0-9a-f]{40} # v[0-9.]+", line);
            }
            Assert.Contains("uses: actions/checkout@", body, StringComparison.Ordinal);
            Assert.Contains("uses: actions/setup-node@", body, StringComparison.Ordinal);
            Assert.Contains("uses: actions/download-artifact@", body, StringComparison.Ordinal);
        }
    }

    [Fact]
    [Trait("Feature", "ReleaseMetadata")]
    public void ReleaseValidation_ExecutesBothArm64PackagesBeforeCreatingTag()
    {
        var workflow = File.ReadAllText(ReleaseWorkflow);
        var nativeTests = ExtractWorkflowJob(workflow, "verify-arm64");

        Assert.Contains("runs-on: windows-11-arm", nativeTests, StringComparison.Ordinal);
        Assert.Contains("architecture: arm64", nativeTests, StringComparison.Ordinal);
        Assert.Contains("Test-NpmPackages.ps1", nativeTests, StringComparison.Ordinal);
        Assert.Contains("@('Cli', 'McpServer')", nativeTests, StringComparison.Ordinal);
        Assert.Contains("-Architecture arm64", nativeTests, StringComparison.Ordinal);
        Assert.Contains("verify-arm64", ExtractWorkflowJob(workflow, "create-tag"), StringComparison.Ordinal);
    }

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
    public void McpRegistryPublication_UsesOnlyExactExistingRelease()
    {
        var release = File.ReadAllText(ReleaseWorkflow);
        var registry = File.ReadAllText(McpRegistryWorkflow);
        var publishMcpRegistry = ExtractWorkflowJob(release, "publish-mcp-registry");

        Assert.Contains("uses: ./.github/workflows/publish-mcp-registry.yml", publishMcpRegistry, StringComparison.Ordinal);
        Assert.Contains("needs: [version, create-tag, create-release, publish]", publishMcpRegistry, StringComparison.Ordinal);
        Assert.DoesNotContain("workflow_dispatch:", registry, StringComparison.Ordinal);
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
    public async Task McpRegistryPublication_RequiresTagCommitOnMain(bool tagCommitOnMain)
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
                function Invoke-FixtureGit {
                    git @args
                    if ($LASTEXITCODE -ne 0) { throw "Fixture Git command failed: git $args" }
                }
                Invoke-FixtureGit init --bare '{{remote.Replace("'", "''", StringComparison.Ordinal)}}'
                Invoke-FixtureGit init -b main '{{repository.Replace("'", "''", StringComparison.Ordinal)}}'
                Set-Location '{{repository.Replace("'", "''", StringComparison.Ordinal)}}'
                Invoke-FixtureGit config user.name fixture
                Invoke-FixtureGit config user.email fixture@example.test
                Set-Content release.txt base
                Invoke-FixtureGit add release.txt
                Invoke-FixtureGit commit -m base
                Invoke-FixtureGit remote add origin '{{remote.Replace("'", "''", StringComparison.Ordinal)}}'
                Invoke-FixtureGit push -u origin main
                if (-not ${{tagCommitOnMain.ToString().ToLowerInvariant()}}) {
                    Invoke-FixtureGit checkout -b unmerged
                    Set-Content release.txt unmerged
                    Invoke-FixtureGit commit -am unmerged
                }
                Invoke-FixtureGit tag v1.2.3
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

    [Fact]
    [Trait("Feature", "ReleaseMetadata")]
    public async Task GitFixtures_IgnoreInheritedHookRepositoryAndIndex()
    {
        var sandbox = CreateSandbox();
        try
        {
            var foreign = Path.Combine(sandbox, "foreign");
            var repository = Path.Combine(sandbox, "fixture");
            var setup = Path.Combine(sandbox, "setup.ps1");
            File.WriteAllText(setup, $$"""
                $ErrorActionPreference = 'Stop'
                $PSNativeCommandUseErrorActionPreference = $true
                git init -b main '{{foreign.Replace("'", "''", StringComparison.Ordinal)}}'
                Set-Location '{{foreign.Replace("'", "''", StringComparison.Ordinal)}}'
                git config user.name original
                git config user.email original@example.test
                Set-Content sentinel.txt untouched
                git add sentinel.txt
                git commit -m sentinel
                """);
            var prepared = await RunPowerShellScriptAsync(setup, [], sandbox);
            Assert.True(prepared.ExitCode == 0, prepared.CombinedOutput);
            var foreignConfig = Path.Combine(foreign, ".git", "config");
            var foreignIndex = Path.Combine(foreign, ".git", "index");
            var configBefore = File.ReadAllBytes(foreignConfig);
            var indexBefore = File.ReadAllBytes(foreignIndex);
            var headBefore = File.ReadAllText(Path.Combine(foreign, ".git", "HEAD"));
            var tipPath = Path.Combine(foreign, ".git", "refs", "heads", "main");
            var tipBefore = File.ReadAllText(tipPath);
            var runner = Path.Combine(sandbox, "fixture.ps1");
            File.WriteAllText(runner, $$"""
                $ErrorActionPreference = 'Stop'
                $PSNativeCommandUseErrorActionPreference = $true
                git init -b main '{{repository.Replace("'", "''", StringComparison.Ordinal)}}'
                Set-Location '{{repository.Replace("'", "''", StringComparison.Ordinal)}}'
                git config user.name fixture
                git config user.email fixture@example.test
                Set-Content release.txt fixture
                git add release.txt
                git commit -m fixture
                git rev-parse --show-toplevel
                $PSNativeCommandUseErrorActionPreference = $false
                $injected = git config --get fixture.injected
                if ($LASTEXITCODE -ne 1) { throw 'Fixture inherited injected Git configuration.' }
                $global:LASTEXITCODE = 0
                """);
            var result = await RunPowerShellScriptAsync(runner, [], sandbox,
                new Dictionary<string, string>
                {
                    ["GIT_DIR"] = Path.Combine(foreign, ".git"),
                    ["GIT_WORK_TREE"] = foreign,
                    ["GIT_INDEX_FILE"] = foreignIndex,
                    ["GIT_CONFIG_COUNT"] = "1",
                    ["GIT_CONFIG_KEY_0"] = "fixture.injected",
                    ["GIT_CONFIG_VALUE_0"] = "foreign"
                });

            Assert.True(result.ExitCode == 0, result.CombinedOutput);
            Assert.Contains(repository.Replace('\\', '/'), result.Stdout, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(configBefore, File.ReadAllBytes(foreignConfig));
            Assert.Equal(indexBefore, File.ReadAllBytes(foreignIndex));
            Assert.Equal(headBefore, File.ReadAllText(Path.Combine(foreign, ".git", "HEAD")));
            Assert.Equal(tipBefore, File.ReadAllText(tipPath));
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

    [Fact]
    [Trait("Feature", "ReleaseMetadata")]
    public void PluginPublication_ReleaseChainNeverCallsManualListingUpdater()
    {
        var workflows = Path.Combine(RepoRoot, ".github", "workflows");
        var release = File.ReadAllText(Path.Combine(workflows, "release.yml"));
        var plugins = File.ReadAllText(Path.Combine(workflows, "publish-plugins.yml"));
        var caller = ExtractWorkflowJob(release, "publish-plugins");
        Assert.Contains("uses: ./.github/workflows/publish-plugins.yml", caller, StringComparison.Ordinal);
        var granted = ExtractWorkflowPermissions(caller, 4);
        Assert.NotNull(granted);
        Assert.Equal(["contents"], granted.Keys);
        Assert.Equal("read", granted["contents"]);

        foreach (var workflow in new[] { release, plugins })
        {
            Assert.DoesNotContain("update-awesome-copilot", workflow, StringComparison.Ordinal);
            Assert.DoesNotContain("AWESOME_COPILOT", workflow, StringComparison.Ordinal);
            Assert.DoesNotContain("COPILOT_GITHUB_TOKEN", workflow, StringComparison.Ordinal);
        }
        Assert.Equal(["contents"], ExtractWorkflowPermissions(
            plugins[..plugins.IndexOf("\njobs:", StringComparison.Ordinal)], 0)!.Keys);
        Assert.Null(ExtractWorkflowPermissions(ExtractWorkflowJob(plugins, "resolve"), 4));
        Assert.Null(ExtractWorkflowPermissions(ExtractWorkflowJob(plugins, "publish"), 4));

        var updater = File.ReadAllText(Path.Combine(workflows, "update-awesome-copilot.md"));
        var trigger = updater[..updater.IndexOf("\npermissions:", StringComparison.Ordinal)];
        Assert.DoesNotContain("workflow_call:", trigger, StringComparison.Ordinal);
        Assert.Contains("workflow_dispatch:", trigger, StringComparison.Ordinal);
        Assert.Contains("published_tag:", trigger, StringComparison.Ordinal);
        Assert.Contains("preview:", trigger, StringComparison.Ordinal);
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
            var scenario = JsonSerializer.Serialize(new
            {
                publishedVersion,
                tagExists,
                manualRepair,
                payloadVersion,
                syncFails,
                succeeds
            });
            var script = Path.Combine(RepoRoot, "tests", "ExcelMcp.Packaging.Tests", "PluginPublication.test.mjs");
            var runner = Path.Combine(sandbox, "run.ps1");
            File.WriteAllText(runner, $$"""
                $env:PLUGIN_PUBLICATION_SCENARIO='{{scenario}}'
                node --test --test-name-pattern 'legacy publication guard scenario' '{{script.Replace("'", "''", StringComparison.Ordinal)}}'
                exit $LASTEXITCODE
                """);
            var result = await RunPowerShellScriptAsync(runner, [], sandbox);
            Assert.True(result.ExitCode == 0, result.CombinedOutput);
            Assert.Contains("legacy publication guard scenario", result.Stdout, StringComparison.Ordinal);
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
                @"(?<mcpTools>\d+) MCP tools across (?<tools>\d+) feature areas, with (?<operations>\d+) operations");
            Assert.True(headline.Success);
            var canonicalTools = int.Parse(headline.Groups["tools"].Value, System.Globalization.CultureInfo.InvariantCulture);
            var canonicalMcpTools = int.Parse(headline.Groups["mcpTools"].Value, System.Globalization.CultureInfo.InvariantCulture);
            var canonicalOperations = int.Parse(
                headline.Groups["operations"].Value,
                System.Globalization.CultureInfo.InvariantCulture);
            var canonicalHeadline = $"{canonicalMcpTools} MCP tools across {canonicalTools} feature areas, with {canonicalOperations} operations";

            CopyDocumentationCountFiles(sandbox, canonicalTools, canonicalMcpTools, canonicalOperations);
            var readmePath = Path.Combine(sandbox, "README.md");
            var llmOutputsPath = Path.Combine(sandbox, "gh-pages", "sitegen", "llm.py");
            await File.WriteAllTextAsync(
                readmePath,
                (await File.ReadAllTextAsync(readmePath))
                    .Replace(
                        canonicalHeadline,
                        "1 MCP tools across 1 feature areas, with 2 operations",
                        StringComparison.Ordinal)
                    .Replace($"all {canonicalOperations} operations", "all 2 operations", StringComparison.Ordinal));

            var update = await RunPowerShellScriptAsync(
                Path.Combine(sandbox, "scripts", "check-doc-counts.ps1"),
                ["-Update", "-SkipBuild"],
                sandbox);

            Assert.True(update.ExitCode == 0, update.CombinedOutput);
            Assert.Contains(
                canonicalHeadline,
                await File.ReadAllTextAsync(readmePath),
                StringComparison.Ordinal);
            Assert.Contains(
                $"all {canonicalOperations} operations",
                await File.ReadAllTextAsync(readmePath),
                StringComparison.Ordinal);
            var llmOutputsContent = await File.ReadAllTextAsync(llmOutputsPath);
            Assert.Contains("json.loads(read(DOC_COUNTS))", llmOutputsContent, StringComparison.Ordinal);
            Assert.Contains("headline_tools, headline_operations = headline_counts()", llmOutputsContent, StringComparison.Ordinal);
            Assert.Contains("for line in read(page.source).splitlines():", llmOutputsContent, StringComparison.Ordinal);
            Assert.DoesNotMatch(
                @"exposing \d+ tools and \d+ operations",
                llmOutputsContent);

            var docCountsPath = Path.Combine(sandbox, "doc-counts.json");
            Assert.True(File.Exists(docCountsPath), "-Update must generate the single doc-counts.json include file.");
            Assert.Equal(canonicalTools, ReadJsonInt(docCountsPath, "tools"));
            Assert.Equal(canonicalOperations, ReadJsonInt(docCountsPath, "operations"));

            var validation = await RunPowerShellScriptAsync(
                Path.Combine(sandbox, "scripts", "check-doc-counts.ps1"),
                ["-SkipBuild"],
                sandbox);
            Assert.True(validation.ExitCode == 0, validation.CombinedOutput);

            var generatedToolsPath = Path.Combine(sandbox, "src", "ExcelMcp.McpServer",
                "obj", "GeneratedFiles", "Tools.g.cs");
            var generatedTools = await File.ReadAllTextAsync(generatedToolsPath);
            await File.WriteAllTextAsync(generatedToolsPath,
                generatedTools.Replace("[McpServerTool(Name = \"tool-1_read\")]", "", StringComparison.Ordinal));
            var missingReadEndpoint = await RunPowerShellScriptAsync(
                Path.Combine(sandbox, "scripts", "check-doc-counts.ps1"),
                ["-SkipBuild", "-AllowStaleAdvertisedCounts"],
                sandbox);
            Assert.NotEqual(0, missingReadEndpoint.ExitCode);
            Assert.Contains("Missing: [tool-1_read]", missingReadEndpoint.CombinedOutput, StringComparison.Ordinal);
            await File.WriteAllTextAsync(generatedToolsPath, generatedTools);

            await File.WriteAllTextAsync(
                readmePath,
                (await File.ReadAllTextAsync(readmePath))
                    .Replace(
                        canonicalHeadline,
                        $"{canonicalMcpTools} MCP tools across 1 feature areas, with {canonicalOperations} operations",
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
                        $"{canonicalMcpTools} MCP tools across 1 feature areas, with {canonicalOperations} operations",
                        canonicalHeadline,
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
        int canonicalMcpTools,
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
            Path.Combine("gh-pages", "sitegen", "llm.py"),
            Path.Combine(".github", "plugins", "excel-mcp", "README.md"),
            Path.Combine(".github", "plugins", "excel-cli", "README.md"),
            Path.Combine("docs", "INSTALLATION-CLI.md"),
            Path.Combine("docs", "guides", "EXCEL-COM-VS-FILE-PARSERS.md"),
            Path.Combine("docs", "COPILOT-PLUGIN-DISTRIBUTION.md"),
            Path.Combine("src", "ExcelMcp.McpServer", ".mcp", "server.json")
        };

        foreach (var relativePath in relativePaths)
        {
            CopyFile(RepoRoot, sandbox, relativePath);
        }

        WriteFile(
            sandbox,
            Path.Combine("artifacts", "generated-skills", "excel-mcp-report-formatting", "SKILL.md"),
            File.ReadAllText(Path.Combine(RepoRoot, "skills", "excel-mcp-report-formatting", "SKILL.md")));

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
        var toolNames = Enumerable.Range(1, canonicalTools - 1)
            .Select(index => $"[McpServerTool(Name = \"tool-{index}\")]")
            .Concat(Enumerable.Range(1, canonicalMcpTools - canonicalTools)
                .Select(index => $"[McpServerTool(Name = \"tool-{index}_read\")]"))
            .Append("[McpServerTool(Name = \"file\")]");
        WriteFile(
            sandbox,
            Path.Combine("src", "ExcelMcp.McpServer", "obj", "GeneratedFiles", "Tools.g.cs"),
            string.Join(Environment.NewLine, toolNames));
        WriteFile(
            sandbox,
            Path.Combine("src", "ExcelMcp.Core", "obj", "GeneratedFiles", "ExcelMcp.Generators",
                "Sbroenne.ExcelMcp.Generators.ServiceRegistryGenerator", "ServiceRegistry.Contracts.g.cs"),
            string.Join(Environment.NewLine, Enumerable.Range(1, canonicalMcpTools - canonicalTools)
                .Select(index => $"case \"tool-{index}_read\":")));
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

    private static Dictionary<string, string>? ExtractWorkflowPermissions(string body, int indentation)
    {
        var lines = body.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n');
        var prefix = new string(' ', indentation);
        var start = Array.FindIndex(lines, line => line.StartsWith($"{prefix}permissions:", StringComparison.Ordinal));
        if (start < 0)
        {
            return null;
        }
        var permissions = new Dictionary<string, string>(StringComparer.Ordinal);
        if (lines[start] == $"{prefix}permissions: {{}}")
        {
            return permissions;
        }
        Assert.Equal($"{prefix}permissions:", lines[start]);
        foreach (var line in lines.Skip(start + 1).TakeWhile(line => line.StartsWith($"{prefix}  ", StringComparison.Ordinal)))
        {
            Assert.Matches(@"^[a-z-]+: (read|write|none)$", line.Trim());
            var entry = line.Trim().Split(": ", StringSplitOptions.None);
            permissions.Add(entry[0], entry[1]);
        }
        Assert.NotEmpty(permissions);
        return permissions;
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

    private static string[] GetGitRepositoryEnvironmentVariables()
    {
        var info = new ProcessStartInfo("git")
        {
            WorkingDirectory = RepoRoot,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        info.ArgumentList.Add("rev-parse");
        info.ArgumentList.Add("--local-env-vars");
        using var process = Process.Start(info);
        Assert.NotNull(process);
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(30000))
        {
            process.Kill(entireProcessTree: true);
            process.WaitForExit();
            throw new TimeoutException("Git environment discovery exceeded 30 seconds.");
        }
        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException($"Git environment discovery failed: {stderr.GetAwaiter().GetResult()}");
        }
        return stdout.GetAwaiter().GetResult().Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
    }

    private static async Task<ScriptResult> RunPowerShellScriptAsync(
        string scriptPath,
        IReadOnlyList<string> arguments,
        string? workingDirectory = null,
        IReadOnlyDictionary<string, string>? environment = null)
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
        if (environment != null)
        {
            foreach (var (key, value) in environment) { startInfo.Environment[key] = value; }
        }
        foreach (var name in GitRepositoryEnvironmentVariables.Concat(
                     startInfo.Environment.Keys.Where(name =>
                         System.Text.RegularExpressions.Regex.IsMatch(name, @"^GIT_CONFIG_(KEY|VALUE)_\d+$")).ToArray()))
        {
            startInfo.Environment.Remove(name);
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
