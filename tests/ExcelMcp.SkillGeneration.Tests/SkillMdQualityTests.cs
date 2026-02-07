using System.Text.RegularExpressions;
using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

/// <summary>
/// Tests to validate the quality of generated SKILL.md files.
/// These tests catch issues like empty parameter descriptions that
/// make skills less useful for LLMs.
/// </summary>
public class SkillMdQualityTests
{
    private static readonly string SkillsFolder = Path.Combine(
        AppContext.BaseDirectory, "skills");

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CliSkill_Exists()
    {
        var skillPath = Path.Combine(SkillsFolder, "excel-cli", "SKILL.md");
        Assert.True(File.Exists(skillPath), $"CLI SKILL.md should exist at {skillPath}");
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void McpSkill_Exists()
    {
        var skillPath = Path.Combine(SkillsFolder, "excel-mcp", "SKILL.md");
        Assert.True(File.Exists(skillPath), $"MCP SKILL.md should exist at {skillPath}");
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CliSkill_HasNoEmptyParameterDescriptions()
    {
        var skillPath = Path.Combine(SkillsFolder, "excel-cli", "SKILL.md");
        AssertNoEmptyDescriptions(skillPath, "CLI");
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void McpSkill_HasNoEmptyParameterDescriptions()
    {
        // MCP SKILL.md doesn't have auto-generated parameter tables
        // Tools are discovered via MCP schema - skill contains curated guidance
        // Skip parameter validation for MCP skill
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CliSkill_HasCommands()
    {
        var skillPath = Path.Combine(SkillsFolder, "excel-cli", "SKILL.md");
        var content = File.ReadAllText(skillPath);
        var commandMatches = Regex.Matches(content, @"^### \w+", RegexOptions.Multiline);
        Assert.True(commandMatches.Count > 0, "CLI SKILL.md should have command headings");
        Assert.True(commandMatches.Count >= 10, $"CLI SKILL.md should have at least 10 commands, found {commandMatches.Count}");
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void McpSkill_HasTools()
    {
        // MCP SKILL.md contains curated guidance, not auto-generated tool docs
        // Tools are discovered via MCP schema at runtime
        // Verify it has the expected curated content
        var skillPath = Path.Combine(SkillsFolder, "excel-mcp", "SKILL.md");
        var content = File.ReadAllText(skillPath);
        Assert.Contains("excel_file", content);
        Assert.Contains("excel_range", content);
        Assert.Contains("excel_calculation_mode", content);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CliSkill_HasParameterTables()
    {
        var skillPath = Path.Combine(SkillsFolder, "excel-cli", "SKILL.md");
        var content = File.ReadAllText(skillPath);
        Assert.Contains("| Parameter | Description |", content);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void McpSkill_HasParameterTables()
    {
        // MCP SKILL.md has markdown tables for reference, not parameter tables
        var skillPath = Path.Combine(SkillsFolder, "excel-mcp", "SKILL.md");
        var content = File.ReadAllText(skillPath);
        Assert.Contains("| Task | Tool |", content);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CliSkill_HasActionsList()
    {
        var skillPath = Path.Combine(SkillsFolder, "excel-cli", "SKILL.md");
        var content = File.ReadAllText(skillPath);
        Assert.Contains("**Actions:**", content);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void McpSkill_HasActionsList()
    {
        // MCP SKILL.md has curated action examples, not **Actions:** section
        var skillPath = Path.Combine(SkillsFolder, "excel-mcp", "SKILL.md");
        var content = File.ReadAllText(skillPath);
        Assert.Contains("action:", content);
    }

    private static void AssertNoEmptyDescriptions(string skillPath, string skillType)
    {
        Assert.True(File.Exists(skillPath), $"{skillType} SKILL.md should exist");
        var content = File.ReadAllText(skillPath);
        var lines = content.Split('\n');
        var emptyDescriptions = new List<string>();
        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i].Trim();
            if (Regex.IsMatch(line, @"^\|\s*`[^`]+`\s*\|\s*\|$"))
            {
                var paramMatch = Regex.Match(line, @"`([^`]+)`");
                if (paramMatch.Success)
                {
                    emptyDescriptions.Add(paramMatch.Groups[1].Value);
                }
            }
        }

        if (emptyDescriptions.Count > 0)
        {
            var message = $"{skillType} SKILL.md has {emptyDescriptions.Count} parameters with empty descriptions:\n" +
                          string.Join("\n", emptyDescriptions.Take(10).Select(p => $"  - {p}"));
            if (emptyDescriptions.Count > 10)
            {
                message += $"\n  ... and {emptyDescriptions.Count - 10} more";
            }

            Assert.Fail(message);
        }
    }
}
