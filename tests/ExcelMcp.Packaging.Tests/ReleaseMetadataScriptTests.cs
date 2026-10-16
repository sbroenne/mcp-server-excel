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

    [Theory]
    [InlineData("absent", false, true)]
    [InlineData("draft-later-page", false, true)]
    [InlineData("draft-delayed", false, true)]
    [InlineData("published", false, true)]
    [InlineData("missing", false, false)]
    [InlineData("missing", true, true)]
    [InlineData("immutable-missing", true, false)]
    [InlineData("mismatch", true, false)]
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
                "ExcelMcp-CLI-1.2.3-windows.zip", "ExcelMcp-CLI-1.2.3-windows-arm64.zip",
                "ExcelMcp-MCP-Server-1.2.3-windows.zip", "ExcelMcp-MCP-Server-1.2.3-windows-arm64.zip",
                "excel-plugins-v1.2.3.zip",
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
                $global:releaseFixtureDraftLookups = 0
                $global:releaseFixtureCalls = [Collections.Generic.List[string]]::new()
                function global:Start-Sleep {
                    param([int]$Seconds)
                    $global:releaseFixtureCalls.Add("Start-Sleep -Seconds $Seconds")
                }
                function global:gh {
                    param([Parameter(ValueFromRemainingArguments)][string[]]$Arguments)
                    $global:releaseFixtureCalls.Add(($Arguments -join ' '))
                    $global:LASTEXITCODE = 0
                    if ($Arguments[0] -eq 'api') {
                        $isList = $Arguments[1] -eq 'repos/owner/repo/releases'
                        $isAbsent = '{{mode}}' -in @('absent', 'draft-delayed') -and -not $global:releaseFixtureCreated
                        $isDraft = '{{mode}}' -in @('absent', 'draft-later-page', 'draft-delayed') `
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
                            if ($isAbsent) { '[[]]'; return }
                            if ('{{mode}}' -eq 'draft-delayed' -and $global:releaseFixtureCreated) {
                                $global:releaseFixtureDraftLookups++
                                if ($global:releaseFixtureDraftLookups -le 2) { '[[]]'; return }
                            }
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
            if (mode is "absent" or "draft-later-page" or "draft-delayed")
            {
                using var response = JsonDocument.Parse(File.ReadAllText(Path.Combine(sandbox, "list-response.json")));
                var pages = response.RootElement;
                Assert.Equal(JsonValueKind.Array, pages.ValueKind);
                Assert.All(pages.EnumerateArray(), page =>
                {
                    Assert.Equal(JsonValueKind.Array, page.ValueKind);
                    Assert.Equal(JsonValueKind.Object, Assert.Single(page.EnumerateArray()).ValueKind);
                });
                Assert.Equal(mode == "draft-later-page" ? 2 : 1, pages.GetArrayLength());
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
                    _ => throw new InvalidOperationException($"Unexpected failure fixture: {mode}")
                };
                Assert.Contains(expectedError, result.CombinedOutput, StringComparison.Ordinal);
            }
            var calls = File.ReadAllText(Path.Combine(sandbox, "calls.json"));
            if (mode is "absent" or "draft-later-page" or "draft-delayed")
            {
                Assert.Contains("release upload", calls, StringComparison.Ordinal);
                Assert.Contains("--clobber", calls, StringComparison.Ordinal);
                Assert.Contains("release edit", calls, StringComparison.Ordinal);
                Assert.Contains("api repos/owner/repo/releases --paginate --slurp", calls, StringComparison.Ordinal);
                Assert.Contains("Published v1.2.3 after verifying every draft asset", result.CombinedOutput, StringComparison.Ordinal);
                if (mode == "draft-delayed")
                {
                    Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(
                        calls, "Start-Sleep -Seconds 2", System.Text.RegularExpressions.RegexOptions.CultureInvariant).Count);
                }
                if (mode is "absent" or "draft-delayed")
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
                Assert.DoesNotContain("api repos/owner/repo/releases --paginate", calls, StringComparison.Ordinal);
                if (mode == "missing" && succeeds)
                {
                    Assert.Contains("release upload", calls, StringComparison.Ordinal);
                }
                else
                {
                    Assert.DoesNotContain("release upload", calls, StringComparison.Ordinal);
                }
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
                Assert.Equal(8, inputs.RootElement.GetProperty("artifacts").GetArrayLength());
                foreach (var artifact in inputs.RootElement.GetProperty("artifacts").EnumerateArray())
                {
                    var name = artifact.GetProperty("name").GetString()!;
                    Assert.Equal(
                        Convert.ToHexStringLower(SHA256.HashData(File.ReadAllBytes(Path.Combine(artifacts, name)))),
                        artifact.GetProperty("sha256").GetString());
                }
                var checksumLines = File.ReadAllLines(Path.Combine(sandbox, "publish", "SHA256SUMS"));
                Assert.Equal(10, checksumLines.Length);
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

    [Theory]
    [InlineData(true, true, false)]
    [InlineData(true, false, false)]
    [InlineData(true, true, true)]
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

    [Theory]
    [InlineData("1.2.2", false, "1.2.3", true)]
    [InlineData("1.2.3", true, "1.2.3", true)]
    [InlineData("1.2.4", false, "1.2.3", false)]
    [InlineData("1.2.2", false, "9.9.9", false)]
    [Trait("Feature", "ReleaseMetadata")]
    public async Task PluginPublication_ExercisesVersionRetryAndFailureDecisions(
        string publishedVersion, bool tagExists, string payloadVersion, bool succeeds)
    {
        var sandbox = CreateSandbox();
        try
        {
            var scenario = JsonSerializer.Serialize(new
            {
                publishedVersion,
                tagExists,
                payloadVersion,
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
    public async Task BuildChangelog_FailsWhenPendingFragmentsProduceNoEntries()
    {
        var sandbox = CreateSandbox();
        try
        {
            CopyReleaseMetadataFiles(sandbox);
            CopyFile(RepoRoot, sandbox, "CHANGELOG.md");
            CopyFile(RepoRoot, sandbox, Path.Combine(".changeset", "config.json"));
            WriteFile(
                sandbox,
                Path.Combine(".changeset", "pending-release-note.md"),
                """
                ---
                "excelmcp": patch
                ---

                A pending release note.
                """);

            var fakeNpxDirectory = Path.Combine(sandbox, "tools");
            Directory.CreateDirectory(fakeNpxDirectory);
            await File.WriteAllTextAsync(
                Path.Combine(fakeNpxDirectory, "npx.cmd"),
                "@echo off\r\ndel \"%~dp0..\\.changeset\\pending-release-note.md\"\r\nexit /b 0\r\n");

            var path = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
            var result = await RunPowerShellScriptAsync(
                BuildChangelogScript,
                ["-Version", "9.8.7", "-RepoRoot", sandbox],
                sandbox,
                new Dictionary<string, string>
                {
                    ["PATH"] = fakeNpxDirectory + Path.PathSeparator + path
                });

            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("pending-release-note.md", result.CombinedOutput, StringComparison.Ordinal);
            Assert.False(File.Exists(Path.Combine(sandbox, "release_notes_body.md")));
        }
        finally
        {
            Directory.Delete(sandbox, recursive: true);
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "ReleaseMetadata")]
    public void DocumentationCounts_WebsiteUsesReviewedMetadataAndCanonicalFeatureSections()
    {
        var content = File.ReadAllText(Path.Combine(RepoRoot, "gh-pages", "sitegen", "llm.py"));
        Assert.Contains("json.loads(read(DOC_COUNTS))", content, StringComparison.Ordinal);
        Assert.Contains("headline_tools, headline_operations = headline_counts()", content, StringComparison.Ordinal);
        Assert.Contains("for line in read(page.source).splitlines():", content, StringComparison.Ordinal);
        Assert.DoesNotMatch(@"exposing \d+ tools and \d+ operations", content);
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
    public async Task ReleaseNotes_UsePlainHeadingsAndThankOnlyOutsideContributors()
    {
        var sandbox = CreateSandbox();
        try
        {
            var notesPath = Path.Combine(sandbox, "release_notes_body.md");
            await File.WriteAllTextAsync(
                notesPath,
                """
                ### Major Changes

                - [#7](https://github.com/sbroenne/mcp-server-excel/pull/7) [`abc1234`](https://github.com/sbroenne/mcp-server-excel/commit/abc1234def) Thanks [@sbroenne](https://github.com/sbroenne)! - Removed an old action.

                ### Minor Changes

                - [#9](https://github.com/sbroenne/mcp-server-excel/pull/9) [`1234567`](https://github.com/sbroenne/mcp-server-excel/commit/1234567abc) Thanks [@sbroenne](https://github.com/sbroenne)! - Add PivotTable sorting.

                  Second paragraph stays.

                ### Patch Changes

                - [#8](https://github.com/sbroenne/mcp-server-excel/pull/8) [`7654321`](https://github.com/sbroenne/mcp-server-excel/commit/7654321abc) Thanks [@helper](https://github.com/helper)! - Faster CLI startup.
                """);

            var result = await RunPowerShellScriptAsync(
                Path.Combine(RepoRoot, "scripts", "Format-ReleaseNotes.ps1"),
                ["-Path", notesPath]);

            Assert.True(result.ExitCode == 0, result.CombinedOutput);
            var notes = (await File.ReadAllTextAsync(notesPath)).ReplaceLineEndings("\n");
            Assert.Equal(
                """
                #### Breaking changes

                - Removed an old action. ([#7](https://github.com/sbroenne/mcp-server-excel/pull/7))

                #### New features

                - Add PivotTable sorting. ([#9](https://github.com/sbroenne/mcp-server-excel/pull/9))

                  Second paragraph stays.

                #### Fixes and improvements

                - Faster CLI startup. ([#8](https://github.com/sbroenne/mcp-server-excel/pull/8)) Thanks [@helper](https://github.com/helper)!
                """.ReplaceLineEndings("\n"),
                notes);

            result = await RunPowerShellScriptAsync(
                Path.Combine(RepoRoot, "scripts", "Format-ReleaseNotes.ps1"),
                ["-Path", notesPath]);
            Assert.True(result.ExitCode == 0, result.CombinedOutput);
            Assert.Equal(notes, (await File.ReadAllTextAsync(notesPath)).ReplaceLineEndings("\n"));
        }
        finally
        {
            Directory.Delete(sandbox, recursive: true);
        }
    }

    [Theory]
    [InlineData("helper,sbroenne,contributor", "helper,contributor")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "ReleaseMetadata")]
    public async Task ReleaseNotes_PreservesMultipleOutsideContributorCredits(string authors, string remainingAuthors)
    {
        var sandbox = CreateSandbox();
        try
        {
            var notesPath = Path.Combine(sandbox, "notes.md");
            var credits = string.Join(", ", authors.Split(',').Select(author =>
                $"[@{author}](https://github.com/{author})"));
            await File.WriteAllTextAsync(notesPath,
                $"- [#9](https://github.com/sbroenne/mcp-server-excel/pull/9) [`abc1234`](https://github.com/sbroenne/mcp-server-excel/commit/abc1234) Thanks {credits}! - Add sorting.");
            var result = await RunPowerShellScriptAsync(
                Path.Combine(RepoRoot, "scripts", "Format-ReleaseNotes.ps1"),
                ["-Path", notesPath]);
            Assert.True(result.ExitCode == 0, result.CombinedOutput);
            var expected = "- Add sorting. ([#9](https://github.com/sbroenne/mcp-server-excel/pull/9))";
            if (remainingAuthors.Length > 0)
            {
                expected += " Thanks " + string.Join(", ", remainingAuthors.Split(',').Select(author =>
                    $"[@{author}](https://github.com/{author})")) + "!";
            }
            Assert.Equal(expected, await File.ReadAllTextAsync(notesPath));
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
