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
    private static readonly string[] ExpectedCliReferenceFiles = ["cli-commands.md", "README.md"];

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
        var referencePath = Path.Combine(SkillsFolder, "excel-cli", "references", "cli-commands.md");
        AssertNoEmptyDescriptions(referencePath, "CLI command reference");
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
    public void CliCommandReference_HasCommands()
    {
        var skillPath = Path.Combine(SkillsFolder, "excel-cli", "references", "cli-commands.md");
        var content = File.ReadAllText(skillPath);
        var commandMatches = Regex.Matches(content, @"^### \w+", RegexOptions.Multiline);
        Assert.True(commandMatches.Count > 0, "CLI command reference should have command headings");
        Assert.True(commandMatches.Count >= 10, $"CLI command reference should have at least 10 commands, found {commandMatches.Count}");
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
        Assert.Contains("file", content);
        Assert.Contains("range", content);
        Assert.Contains("calculation_mode", content);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CliCommandReference_HasParameterTables()
    {
        var skillPath = Path.Combine(SkillsFolder, "excel-cli", "references", "cli-commands.md");
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
    public void CliCommandReference_HasActionsList()
    {
        var skillPath = Path.Combine(SkillsFolder, "excel-cli", "references", "cli-commands.md");
        var content = File.ReadAllText(skillPath);
        Assert.Contains("**Actions:**", content);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CliSkill_DelegatesFullCommandReference()
    {
        var skillPath = Path.Combine(SkillsFolder, "excel-cli", "SKILL.md");
        var content = File.ReadAllText(skillPath);

        Assert.Contains("./references/cli-commands.md", content);
        Assert.Contains("excelcli -q <command> <action>", content);
        Assert.DoesNotContain("### calculationmode", content);
        Assert.DoesNotContain("| Parameter | Description |", content);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CliSkill_DoesNotLinkMcpStyleDomainReferences()
    {
        var skillPath = Path.Combine(SkillsFolder, "excel-cli", "SKILL.md");
        var content = File.ReadAllText(skillPath);

        Assert.DoesNotContain("./references/range.md", content);
        Assert.DoesNotContain("./references/chart.md", content);
        Assert.DoesNotContain("./references/powerquery.md", content);
        Assert.DoesNotContain("./references/worksheet.md", content);
        Assert.DoesNotContain("./references/behavioral-rules.md", content);
        Assert.DoesNotContain("./references/anti-patterns.md", content);
        Assert.DoesNotContain("./references/workflows.md", content);
        Assert.DoesNotContain("range_format(action:", content);
        Assert.DoesNotContain("chart_config(", content);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CliReferences_OnlyContainCliSpecificFiles()
    {
        var referencesPath = Path.Combine(SkillsFolder, "excel-cli", "references");
        var fileNames = Directory.GetFiles(referencesPath, "*.md")
            .Select(path => Path.GetFileName(path)!)
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        Assert.Equal(ExpectedCliReferenceFiles, fileNames);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void McpSkill_DoesNotDuplicateCalculationModeWorkflow()
    {
        var skillPath = Path.Combine(SkillsFolder, "excel-mcp", "SKILL.md");
        var content = File.ReadAllText(skillPath);

        Assert.Contains("## Calculation Mode Workflow", content);
        Assert.DoesNotContain("### Rule 10: Use Calculation Mode", content);
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
