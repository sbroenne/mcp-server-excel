using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

public sealed partial class AutomationSafetyTests
{
    [Theory]
    [InlineData("success", "Installed ponytail-review release v2.0.0.", 2)]
    [InlineData("release-failure", "Ponytail release lookup failed", 0)]
    [InlineData("empty-release", "The latest Ponytail release has no tag", 0)]
    [InlineData("invalid-revision", "The latest Ponytail release has no valid commit SHA", 0)]
    [InlineData("download-failure", "Ponytail archive download failed", 1)]
    [InlineData("missing-skill", "The Ponytail release archive has no review SKILL.md", 1)]
    [InlineData("copy-failure", "Ponytail skill copy failed", 1)]
    [InlineData("invalid-layout", "Unexpected Ponytail release archive layout", 1)]
    public async Task PonytailSetup_RefreshesOnlyReviewSkillAndReportsFailures(
        string scenario, string expectedOutput, int downloadCount)
    {
        var root = NewSandbox();
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
            if (scenario != "missing-skill")
            {
                File.WriteAllText(Path.Combine(sourceSkill, "SKILL.md"), "released review skill v2");
            }
            var nextArchive = Path.Combine(root, "next-release.zip");
            System.IO.Compression.ZipFile.CreateFromDirectory(Path.GetDirectoryName(archiveSource)!, nextArchive);
            var calls = Path.Combine(root, "calls.txt");
            var downloads = Path.Combine(root, "downloads.txt");
            var target = Quote(Path.Combine(RepoRoot, "scripts", "Install-CopilotPonytailReview.ps1"));
            var result = await RunAsync(root, $$"""
                $global:releaseTag = 'v1.0.0'
                $env:GH_TOKEN = 'fixture-token'
                function gh { throw 'Ponytail setup must not require GitHub CLI.' }
                function Invoke-RestMethod {
                    param($Uri, $Headers, $TimeoutSec)
                    Add-Content -LiteralPath '{{Quote(calls)}}' -Value $Uri
                    if ($Headers.Authorization -ne 'Bearer fixture-token') { throw 'Expected scoped GitHub authentication.' }
                    if ($Uri.EndsWith('/releases/latest')) {
                        if ('{{scenario}}' -eq 'release-failure') { throw 'Ponytail release lookup failed.' }
                        if ('{{scenario}}' -eq 'empty-release') { return @{ tag_name = $null } }
                        return @{ tag_name = $global:releaseTag }
                    }
                    if ('{{scenario}}' -eq 'invalid-revision') { return @{ sha = 'invalid' } }
                    $sha = if ($global:releaseTag -eq 'v1.0.0') {
                        '0123456789012345678901234567890123456789'
                    } else { '1123456789012345678901234567890123456789' }
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
                    $sourceArchive = if ($Uri.EndsWith('/1123456789012345678901234567890123456789')) {
                        '{{Quote(nextArchive)}}'
                    } else { '{{Quote(archive)}}' }
                    [IO.File]::Copy($sourceArchive, $OutFile)
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
                $global:releaseTag = 'v2.0.0'
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
                    call == "https://api.github.com/repos/DietrichGebert/ponytail/zipball/0123456789012345678901234567890123456789" ||
                    call == "https://api.github.com/repos/DietrichGebert/ponytail/zipball/1123456789012345678901234567890123456789",
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
                    "https://api.github.com/repos/DietrichGebert/ponytail/commits/v1.0.0",
                    recorded);
                Assert.Contains(
                    "https://api.github.com/repos/DietrichGebert/ponytail/commits/v2.0.0",
                    recorded);
                Assert.Contains(
                    "https://api.github.com/repos/DietrichGebert/ponytail/zipball/1123456789012345678901234567890123456789",
                    recorded);
                Assert.Equal("released review skill v2", File.ReadAllText(installed));
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
}
