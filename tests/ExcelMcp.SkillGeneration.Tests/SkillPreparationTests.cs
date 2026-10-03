using System.Text.RegularExpressions;
using Sbroenne.ExcelMcp.Tests.Infrastructure;
using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

[Collection("GeneratedAssets")]
[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "SkillGeneration")]
public sealed class SkillPreparationTests
{
    private static readonly string RepoRoot = FindRepoRoot();

    [Fact]
    public void PreparedSkills_ContainOnlyScopedFormattingGuidance()
    {
        Assert.Equal(
            ["excel-cli-report-formatting", "excel-mcp-report-formatting"],
            Directory.GetDirectories(GeneratedAssetsFixture.SkillsDirectory)
                .Select(path => new DirectoryInfo(path).Name).Order(StringComparer.Ordinal).ToArray());
        foreach (var skill in Directory.GetDirectories(GeneratedAssetsFixture.SkillsDirectory))
        {
            Assert.Equal(
                ["report-formatting.md"],
                Directory.GetFiles(Path.Combine(skill, "references"))
                    .Select(path => new FileInfo(path).Name).Order(StringComparer.Ordinal).ToArray());
            var name = Path.GetFileName(skill);
            var content = File.ReadAllText(Path.Combine(skill, "SKILL.md"));
            var header = Regex.Match(content, @"\A---\r?\n(?<header>.*?)\r?\n---(?:\r?\n|\z)", RegexOptions.Singleline);
            Assert.True(header.Success, "Skill metadata header is missing.");
            Assert.Matches($@"(?m)^name:\s*{Regex.Escape(name)}\s*$", header.Groups["header"].Value);
            var description = Regex.Match(header.Groups["header"].Value,
                @"(?m)^description: >-\r?\n(?<text>(?:[ \t]+[^\r\n]+(?:\r?\n|$))+)");
            Assert.True(description.Success, "Authored folded description is missing.");
            Assert.False(string.IsNullOrWhiteSpace(description.Groups["text"].Value));
            Assert.True(File.Exists(Path.Combine(RepoRoot, "docs", "reference", "powerquery.md")));
            Assert.False(Directory.Exists(Path.Combine(RepoRoot, "skills", "shared")));
        }
    }

    [Fact]
    public void PackagedReferences_AreReachableFromEachSkill()
    {
        foreach (var skill in new[] { "excel-cli-report-formatting", "excel-mcp-report-formatting" })
        {
            var root = Path.Combine(GeneratedAssetsFixture.SkillsDirectory, skill);
            var pending = new Stack<string>();
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            pending.Push(Path.Combine(root, "SKILL.md"));
            while (pending.TryPop(out var path))
            {
                path = Path.GetFullPath(path);
                if (!visited.Add(path))
                    continue;
                Assert.True(File.Exists(path), $"Missing linked file: {path}");
                foreach (Match match in Regex.Matches(File.ReadAllText(path), @"\]\(([^)]+)\)"))
                {
                    var target = match.Groups[1].Value.Split('#')[0];
                    if (target.Length > 0 && !target.Contains("://", StringComparison.Ordinal)
                        && target.EndsWith(".md", StringComparison.Ordinal))
                        pending.Push(Path.Combine(Path.GetDirectoryName(path)!, target));
                }
            }
            foreach (var reference in Directory.GetFiles(Path.Combine(root, "references"), "*.md", SearchOption.AllDirectories))
                Assert.Contains(Path.GetFullPath(reference), visited);
        }
    }

    private static string FindRepoRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Sbroenne.ExcelMcp.sln"))) { return directory.FullName; }
            directory = directory.Parent;
        }
        throw new DirectoryNotFoundException("Could not locate repository root.");
    }
}
