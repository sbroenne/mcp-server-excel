using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

public sealed partial class AutomationSafetyTests
{
    [Theory]
    [InlineData("success", "Installed ponytail-review release v2.0.0.", 2)]
    [InlineData("release-failure", "Could not resolve the latest Ponytail release", 0)]
    [InlineData("empty-release", "The latest Ponytail release has no tag", 0)]
    [InlineData("install-failure", "Ponytail review skill installation failed", 1)]
    [InlineData("missing-skill", "Ponytail review skill installation did not create SKILL.md", 1)]
    public async Task PonytailSetup_RefreshesOnlyReviewSkillAndReportsFailures(
        string scenario, string expectedOutput, int installCount)
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
            var calls = Path.Combine(root, "calls.txt");
            var target = Quote(Path.Combine(RepoRoot, "scripts", "Install-CopilotPonytailReview.ps1"));
            var result = await RunAsync(root, $$"""
                $global:releaseTag = 'v1.0.0'
                function gh {
                    $call = $args -join ' '
                    Add-Content -LiteralPath '{{Quote(calls)}}' -Value $call
                    $global:LASTEXITCODE = 0
                    if ($args[0] -eq 'api') {
                        if ('{{scenario}}' -eq 'release-failure') { $global:LASTEXITCODE = 23; return }
                        if ('{{scenario}}' -eq 'empty-release') { return }
                        $global:releaseTag
                        return
                    }
                    $null = New-Item -ItemType Directory -Path '{{Quote(Path.GetDirectoryName(installed)!)}}' -Force
                    Set-Content -LiteralPath '{{Quote(installed)}}' -Value 'partial skill'
                    if ('{{scenario}}' -eq 'install-failure') { $global:LASTEXITCODE = 24; return }
                    if ('{{scenario}}' -eq 'missing-skill') {
                        Remove-Item -LiteralPath '{{Quote(installed)}}'
                        return
                    }
                    Set-Content -LiteralPath '{{Quote(installed)}}' -Value $global:releaseTag
                }
                & '{{target}}' -SkillsDirectory '{{Quote(skills)}}'
                $global:releaseTag = 'v2.0.0'
                & '{{target}}' -SkillsDirectory '{{Quote(skills)}}'
                """);
            Assert.Equal(scenario == "success", result.ExitCode == 0);
            Assert.Contains(expectedOutput, result.Output, StringComparison.Ordinal);
            var recorded = File.ReadAllLines(calls);
            Assert.Equal(
                installCount,
                recorded.Count(call => call.StartsWith("skill install ", StringComparison.Ordinal)));
            Assert.All(
                recorded.Where(call => call.StartsWith("api ", StringComparison.Ordinal)),
                call => Assert.Equal("api repos/DietrichGebert/ponytail/releases/latest --jq .tag_name", call));
            if (scenario == "success")
            {
                Assert.Contains(
                    $"skill install DietrichGebert/ponytail ponytail-review@v1.0.0 --dir {skills} --force",
                    recorded);
                Assert.Contains(
                    $"skill install DietrichGebert/ponytail ponytail-review@v2.0.0 --dir {skills} --force",
                    recorded);
                Assert.Equal("v2.0.0", File.ReadAllText(installed).Trim());
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
