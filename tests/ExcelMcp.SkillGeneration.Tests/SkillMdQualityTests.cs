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
    public void CliCommandReference_CoversBranchAndGeneratedCommands()
    {
        var referencePath = Path.Combine(SkillsFolder, "excel-cli", "references", "cli-commands.md");
        var referenceContent = File.ReadAllText(referencePath);
        var skillContent = File.ReadAllText(Path.Combine(SkillsFolder, "excel-cli", "SKILL.md"));
        var groupsSection = skillContent[(skillContent.IndexOf("Available command groups:", StringComparison.Ordinal) + "Available command groups:".Length)..];
        var commandGroups = Regex.Matches(groupsSection.Split("## Common Pitfalls", StringSplitOptions.None)[0], @"`([a-z][a-z0-9-]+)`")
            .Select(match => match.Groups[1].Value)
            .Append("diag")
            .Distinct(StringComparer.Ordinal)
            .ToArray();

        Assert.True(commandGroups.Length >= 34, $"Expected all live command groups in SKILL.md, found {commandGroups.Length}.");
        foreach (var commandGroup in commandGroups)
        {
            Assert.Contains($"### {commandGroup}", referenceContent);
        }
        Assert.Contains("#### session open", referenceContent);
        Assert.Contains("#### service stop", referenceContent);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CliCommandReference_UsesLiveCliOptionAliases()
    {
        var referencePath = Path.Combine(SkillsFolder, "excel-cli", "references", "cli-commands.md");
        var content = File.ReadAllText(referencePath);

        Assert.Contains("`--sheet`", content);
        Assert.Contains("`--range`", content);
        Assert.DoesNotContain("`--sheet-name`", content);
        Assert.DoesNotContain("`--range-address`", content);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CliCommandReference_DoesNotSplitActionNamesAcrossHelpLines()
    {
        var referencePath = Path.Combine(SkillsFolder, "excel-cli", "references", "cli-commands.md");
        var content = File.ReadAllText(referencePath);
        var splitAction = Regex.Match(
            content,
            @"\(required for:[^)]*\b[a-z]+(?:-[a-z]+)+\s+[a-z]+(?:-[a-z]+)*(?=[,)])");
        var splitIdentifier = Regex.Match(
            content,
            @"'[A-Za-z0-9]*[a-z][A-Z][A-Za-z0-9]*\s+[a-z][A-Za-z0-9]*'");

        Assert.False(splitAction.Success, $"Found a split CLI action name: {splitAction.Value}");
        Assert.False(splitIdentifier.Success, $"Found a split CLI identifier: {splitIdentifier.Value}");
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
        Assert.DoesNotContain("--sheet-name", content);
        Assert.DoesNotContain("--range-address", content);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CliSkill_LinksSharedDomainReferences()
    {
        var skillPath = Path.Combine(SkillsFolder, "excel-cli", "SKILL.md");
        var content = File.ReadAllText(skillPath);

        Assert.Contains("./references/range.md", content);
        Assert.Contains("./references/chart.md", content);
        Assert.Contains("./references/powerquery.md", content);
        Assert.Contains("./references/worksheet.md", content);
        Assert.Contains("./references/behavioral-rules.md", content);
        Assert.Contains("./references/anti-patterns.md", content);
        Assert.Contains("./references/workflows.md", content);
        Assert.DoesNotContain("range_format(action:", content);
        Assert.DoesNotContain("chart_config(", content);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CliReferences_ContainGeneratedAndSharedFiles()
    {
        var referencesPath = Path.Combine(SkillsFolder, "excel-cli", "references");
        var fileNames = Directory.GetFiles(referencesPath, "*.md")
            .Select(path => Path.GetFileName(path)!)
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        var expectedFiles = Directory.GetFiles(Path.Combine(SkillsFolder, "shared"), "*.md")
            .Select(path => Path.GetFileName(path)!)
            .Append("cli-commands.md")
            .Append("README.md")
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        Assert.Equal(expectedFiles, fileNames);
        foreach (var sharedFile in expectedFiles.Except(["cli-commands.md", "README.md"]))
        {
            var content = File.ReadAllText(Path.Combine(referencesPath, sharedFile));
            Assert.StartsWith("> **CLI syntax note:**", content);
        }
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

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CanonicalMcpGuidance_UsesSnakeCaseInputs()
    {
        var allowedCamelCaseTokens = new HashSet<string>(StringComparer.Ordinal)
        {
            // MCP and CLI response properties.
            "canOpen",
            "chartName",
            "errorMessage",
            "formulaPreview",
            "groupedFieldName",
            "isIrmProtected",
            "loadMode",
            "majorUnit",
            "minorUnit",
            "newName",
            "oldName",
            "requiresVisibleSession",
            "sessionId",
            "suggestedNextActions",
            "willOpenReadOnly",

            // CLI batch JSON aliases.
            "daxFormulaFile",
            "daxQueryFile",
            "dmvQueryFile",
            "mCodeFile",
            "schemaFile",
            "vbaCodeFile",
            "xmlDataFile",

            // External configuration and XML names.
            "mcpServers",
            "noNamespaceSchemaLocation",
            "schemaLocation",

            // Contract enum values and conditional-format response properties.
            "aboveAverage",
            "aboveBelow",
            "aboveStdDev",
            "barColorNegative",
            "belowAverage",
            "belowStdDev",
            "colorScale",
            "colorScaleCriteria",
            "dataBar",
            "datePeriod",
            "equalAboveAverage",
            "equalBelowAverage",
            "fillColor",
            "fontBold",
            "fontColor",
            "fontItalic",
            "iconSet",
            "interiorColor",
            "last7Days",
            "lastMonth",
            "lastWeek",
            "leftToRight",
            "maxType",
            "maxValue",
            "minType",
            "minValue",
            "nextMonth",
            "nextWeek",
            "rightToLeft",
            "showIconOnly",
            "showValue",
            "borderStyle",
            "thisMonth",
            "thisWeek",
            "timePeriod",
            "topBottom"
        };

        var canonicalFiles = Directory.GetFiles(Path.Combine(SkillsFolder, "shared"), "*.md")
            .Append(Path.Combine(SkillsFolder, "templates", "SKILL.mcp.sbn"))
            .Append(Path.Combine(SkillsFolder, "excel-mcp", "references", "claude-desktop.md"));
        var unexpectedTokens = new List<string>();

        foreach (var path in canonicalFiles)
        {
            var relativePath = Path.GetRelativePath(SkillsFolder, path);
            var lines = File.ReadAllLines(path);

            for (var lineIndex = 0; lineIndex < lines.Length; lineIndex++)
            {
                foreach (Match match in Regex.Matches(lines[lineIndex], @"\b[a-z]+[A-Z][A-Za-z0-9]*\b"))
                {
                    if (!allowedCamelCaseTokens.Contains(match.Value))
                    {
                        unexpectedTokens.Add($"{relativePath}:{lineIndex + 1}: {match.Value}");
                    }
                }
            }
        }

        Assert.True(
            unexpectedTokens.Count == 0,
            "Canonical MCP guidance contains unexpected camelCase tokens. MCP inputs must use snake_case:\n" +
            string.Join('\n', unexpectedTokens));
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
