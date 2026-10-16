using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

public sealed partial class AutomationSafetyTests
{
    [Theory]
    [InlineData("success", "Installed ponytail-review release v5.2.0.", 2)]
    [InlineData("release-failure", "Ponytail release lookup failed", 0)]
    [InlineData("empty-release", "The Ponytail release does not match the pinned tag", 0)]
    [InlineData("invalid-revision", "The Ponytail release does not match the pinned commit SHA", 0)]
    [InlineData("moved-tag", "The Ponytail release does not match the pinned commit SHA", 0)]
    [InlineData("download-failure", "Ponytail archive download failed", 1)]
    [InlineData("missing-skill", "The Ponytail release archive has no review SKILL.md", 1)]
    [InlineData("copy-failure", "Ponytail skill copy failed", 1)]
    [InlineData("invalid-layout", "Unexpected Ponytail release archive layout", 1)]
    public async Task PonytailSetup_RefreshesOnlyReviewSkillAndReportsFailures(
        string scenario, string expectedOutput, int downloadCount)
    {
        var root = Directory.CreateDirectory(Path.Combine(
            RepoRoot, "artifacts", "ponytail-tests", Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            var skills = Path.Combine(root, "skills");
            var installed = Path.Combine(skills, "ponytail-review", "SKILL.md");
            Directory.CreateDirectory(Path.GetDirectoryName(installed)!);
            File.WriteAllText(installed, "stale skill");
            var unrelated = Path.Combine(skills, "other-skill", "SKILL.md");
            Directory.CreateDirectory(Path.GetDirectoryName(unrelated)!);
            File.WriteAllText(unrelated, "keep this skill");
            var archiveSource = Path.Combine(root, "archive-source", "ponytail-fixture");
            var sourceSkill = Path.Combine(archiveSource, "skills", "ponytail-review");
            Directory.CreateDirectory(sourceSkill);
            if (scenario != "missing-skill")
            {
                File.WriteAllText(Path.Combine(sourceSkill, "SKILL.md"), "released review skill v1");
            }
            var references = Directory.CreateDirectory(Path.Combine(sourceSkill, "references")).FullName;
            File.WriteAllText(Path.Combine(references, "guide.txt"), "released reference");
            File.WriteAllText(Path.Combine(archiveSource, "LICENSE"), "fixture license");
            Directory.CreateDirectory(Path.Combine(archiveSource, "skills", "ponytail"));
            File.WriteAllText(Path.Combine(archiveSource, "skills", "ponytail", "SKILL.md"), "do not install");
            if (scenario == "invalid-layout")
            {
                Directory.CreateDirectory(Path.Combine(root, "archive-source", "extra-root"));
            }
            var archive = Path.Combine(root, "release.zip");
            System.IO.Compression.ZipFile.CreateFromDirectory(Path.GetDirectoryName(archiveSource)!, archive);
            var calls = Path.Combine(root, "calls.txt");
            var downloads = Path.Combine(root, "downloads.txt");
            var target = Quote(Path.Combine(RepoRoot, "scripts", "Install-CopilotPonytailReview.ps1"));
            var result = await RunAsync(root, $$"""
                $env:GH_TOKEN = 'fixture-token'
                function gh { throw 'Ponytail setup must not require GitHub CLI.' }
                function Invoke-RestMethod {
                    param($Uri, $Headers, $TimeoutSec)
                    Add-Content -LiteralPath '{{Quote(calls)}}' -Value $Uri
                    if ($Headers.Authorization -ne 'Bearer fixture-token') { throw 'Expected scoped GitHub authentication.' }
                    if ($Uri.EndsWith('/releases/latest')) { throw 'Setup must not select an unreviewed release.' }
                    if ($Uri.EndsWith('/releases/tags/v5.2.0')) {
                        if ('{{scenario}}' -eq 'release-failure') { throw 'Ponytail release lookup failed.' }
                        if ('{{scenario}}' -eq 'empty-release') { return @{ tag_name = $null } }
                        return @{ tag_name = 'v5.2.0' }
                    }
                    if ('{{scenario}}' -eq 'invalid-revision') { return @{ sha = 'invalid' } }
                    if ('{{scenario}}' -eq 'moved-tag') { return @{ sha = '1123456789012345678901234567890123456789' } }
                    $sha = '2961b52fc1fe6a5a87c7a959a79512e5c2f3ffe5'
                    return @{ sha = $sha }
                }
                function Invoke-WebRequest {
                    param($Uri, $Headers, $OutFile, $TimeoutSec)
                    Add-Content -LiteralPath '{{Quote(calls)}}' -Value $Uri
                    Add-Content -LiteralPath '{{Quote(downloads)}}' -Value $OutFile
                    if ('{{scenario}}' -eq 'download-failure') {
                        [IO.File]::WriteAllText($OutFile, 'partial archive')
                        throw 'Ponytail archive download failed.'
                    }
                    [IO.File]::Copy('{{Quote(archive)}}', $OutFile)
                }
                function Copy-Item {
                    param($LiteralPath, $Destination, [switch]$Recurse)
                    if ('{{scenario}}' -eq 'copy-failure') {
                        $null = New-Item -ItemType Directory -Path '{{Quote(Path.GetDirectoryName(installed)!)}}' -Force
                        Set-Content -LiteralPath '{{Quote(installed)}}' -Value 'partial skill'
                        throw 'Ponytail skill copy failed.'
                    }
                    Microsoft.PowerShell.Management\Copy-Item -LiteralPath $LiteralPath -Destination $Destination -Recurse:$Recurse
                }
                & '{{target}}' -SkillsDirectory '{{Quote(skills)}}'
                & '{{target}}' -SkillsDirectory '{{Quote(skills)}}'
                """);
            Assert.Equal(scenario == "success", result.ExitCode == 0);
            Assert.Contains(expectedOutput, result.Output, StringComparison.Ordinal);
            var recorded = File.ReadAllLines(calls);
            Assert.Equal(
                downloadCount,
                recorded.Count(call => call.Contains("/zipball/", StringComparison.Ordinal)));
            Assert.All(
                recorded.Where(call => call.Contains("/zipball/", StringComparison.Ordinal)),
                call => Assert.True(
                    call == "https://api.github.com/repos/DietrichGebert/ponytail/zipball/2961b52fc1fe6a5a87c7a959a79512e5c2f3ffe5",
                    call));
            if (File.Exists(downloads))
            {
                Assert.All(
                    File.ReadAllLines(downloads),
                    download => Assert.False(Directory.Exists(Path.GetDirectoryName(download))));
            }
            if (scenario == "success")
            {
                Assert.Contains(
                    "https://api.github.com/repos/DietrichGebert/ponytail/commits/v5.2.0",
                    recorded);
                Assert.Contains(
                    "https://api.github.com/repos/DietrichGebert/ponytail/zipball/2961b52fc1fe6a5a87c7a959a79512e5c2f3ffe5",
                    recorded);
                Assert.Equal("released review skill v1", File.ReadAllText(installed));
                Assert.Equal("released reference", File.ReadAllText(Path.Combine(skills, "ponytail-review", "references", "guide.txt")));
                Assert.Equal("fixture license", File.ReadAllText(Path.Combine(skills, "ponytail-review", "LICENSE")));
                Assert.Equal(2, Directory.GetDirectories(skills).Length);
            }
            else
            {
                Assert.DoesNotContain("Installed ponytail-review release", result.Output, StringComparison.Ordinal);
                Assert.False(Directory.Exists(Path.GetDirectoryName(installed)));
            }
            Assert.Equal("keep this skill", File.ReadAllText(unrelated));
        }
        finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void CopilotSetup_LeavesEvaluationsOnDemandAndPreservesReviewIsolation()
    {
        var app = File.ReadAllText(Path.Combine(RepoRoot, ".github", "github-app.yml"));
        var coding = File.ReadAllText(Path.Combine(RepoRoot, ".github", "workflows", "copilot-setup-steps.yml"));
        var review = File.ReadAllText(Path.Combine(RepoRoot, ".github", "workflows", "copilot-code-review.yml"));
        foreach (var configuration in new[] { app, coding, review })
        {
            Assert.DoesNotContain("uv sync", configuration, StringComparison.Ordinal);
            Assert.DoesNotContain("setup-uv", configuration, StringComparison.Ordinal);
            Assert.DoesNotContain("llm-tests/", configuration, StringComparison.Ordinal);
        }
        Assert.DoesNotContain("triggers:", app, StringComparison.Ordinal);
        Assert.Contains("runs-on: ubuntu-latest", review, StringComparison.Ordinal);
        Assert.Contains("-p:EnableWindowsTargeting=true", review, StringComparison.Ordinal);
        Assert.DoesNotContain("dotnet test", review, StringComparison.Ordinal);
        Assert.Contains("EXCEL_COPILOT_ENABLED", coding, StringComparison.Ordinal);
        Assert.Contains("Assert-ExcelRunnerReadiness", coding, StringComparison.Ordinal);
        foreach (var workflow in new[] { coding, review })
        {
            Assert.Contains("contents: read", workflow, StringComparison.Ordinal);
            Assert.Contains("persist-credentials: false", workflow, StringComparison.Ordinal);
            Assert.Contains("GH_TOKEN: ${{ github.token }}", workflow, StringComparison.Ordinal);
            Assert.Contains("scripts\\Install-CopilotPonytailReview.ps1", workflow, StringComparison.Ordinal);
            Assert.DoesNotContain("continue-on-error", workflow, StringComparison.Ordinal);
        }
    }
}
