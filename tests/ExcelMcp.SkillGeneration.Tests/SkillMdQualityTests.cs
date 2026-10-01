using System.Text.RegularExpressions;
using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

/// <summary>
/// Tests to validate the quality of generated SKILL.md files.
/// These tests catch issues like empty parameter descriptions that
/// make skills less useful for LLMs.
/// </summary>
[Trait("RequiresExcel", "false")]
[Collection("GeneratedAssets")]
public class SkillMdQualityTests
{
    private static string SkillsFolder => GeneratedAssetsFixture.SkillsDirectory;

    [Theory]
    [InlineData("excel-cli", "--overwrite-policy allow")]
    [InlineData("excel-mcp", "overwrite_policy: 'allow'")]
    public void OverwriteGuidance_ExplainsProtectedDefaultAndAuthorizedReplacement(string skill, string option)
    {
        var root = Path.Combine(SkillsFolder, skill);
        var template = File.ReadAllText(Path.Combine(root, "SKILL.md"));
        Assert.Contains(option, template, StringComparison.Ordinal);
        Assert.Contains("automatically retry", template, StringComparison.Ordinal);
        var range = File.ReadAllText(Path.Combine(root, "references", "range.md"));
        Assert.Contains("default to `reject-nonempty`", range.Replace("\r\n", "\n"), StringComparison.Ordinal);
        Assert.Contains("at most 10 cell addresses", range, StringComparison.Ordinal);
        Assert.Contains("Failed inspection also stops", range, StringComparison.Ordinal);
        Assert.Contains("not a", range, StringComparison.Ordinal);
        Assert.Contains("transaction", range, StringComparison.Ordinal);
        Assert.Contains("whole multiples", range, StringComparison.Ordinal);
        Assert.DoesNotContain("Read before overwriting", range, StringComparison.Ordinal);
    }

    [Fact]
    public void McpSessionGuidance_UsesCanonicalReturnedIdentifiers()
    {
        var content = File.ReadAllText(Path.Combine(SkillsFolder, "excel-mcp", "SKILL.md"));
        Assert.Contains("session error", content, StringComparison.Ordinal);
        Assert.Contains("same `session_id` spelling", content, StringComparison.Ordinal);
        Assert.Contains("`sessionId` is not an accepted input", content, StringComparison.Ordinal);
        Assert.DoesNotContain("entries instead contain", content, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("excel-cli")]
    [InlineData("excel-mcp")]
    public void Guidance_ExplainsCalculationOrderingAndDestructiveRecovery(string skill)
    {
        var content = File.ReadAllText(Path.Combine(SkillsFolder, skill, "SKILL.md"));
        Assert.Contains("concurrent requests and", content, StringComparison.Ordinal);
        Assert.Contains("no guaranteed order", content, StringComparison.Ordinal);
        Assert.Contains("manual needs explicit calculation", content, StringComparison.Ordinal);
        Assert.Contains("attempt to restore the prior mode", content, StringComparison.Ordinal);
        Assert.Contains("Restoration can fail without failing the write", content, StringComparison.Ordinal);
        Assert.Contains("what-if data tables, not worksheet Tables", content, StringComparison.Ordinal);
        var recovery = File.ReadAllText(Path.Combine(SkillsFolder, skill, "references", "behavioral-rules.md"));
        Assert.Contains("no tool-level", recovery, StringComparison.Ordinal);
        Assert.Contains("earlier unsaved work", recovery, StringComparison.Ordinal);
        Assert.Contains("cannot be reversed", recovery, StringComparison.Ordinal);
    }

    [Fact]
    public void CliGuidance_UsesCliExamplesInsteadOfMcpCalls()
    {
        foreach (var path in Directory.GetFiles(Path.Combine(SkillsFolder, "excel-cli"), "*.md", SearchOption.AllDirectories))
        {
            var content = File.ReadAllText(path);
            Assert.DoesNotContain("CLI syntax note", content);
            Assert.False(Regex.IsMatch(content,
                @"\b(?:file|worksheet|range(?:_[a-z]+)?|table(?:_[a-z]+)?|chart(?:_[a-z]+)?|slicer|powerquery|calculation_mode|pivottable(?:_[a-z]+)?|window|analysis|drawing|xmlmap|querytable|screenshot|datamodel(?:_[a-z]+)?|conditionalformat)\("),
                $"MCP call in {Path.GetFileName(path)}");
        }
    }

    [Fact]
    public void CliCommandReference_IsSmallIndexWithCommandPages()
    {
        var references = Path.Combine(SkillsFolder, "excel-cli", "references");
        var index = File.ReadAllText(Path.Combine(references, "cli-commands.md"));
        Assert.True(index.Length < 6000, $"CLI index has {index.Length} characters.");
        Assert.Contains("commands/range.md", index);
        Assert.Contains("--values", File.ReadAllText(Path.Combine(references, "commands", "range.md")));
    }

    [Theory]
    [InlineData("excel-cli")]
    [InlineData("excel-mcp")]
    public void EveryReference_IsReachableFromSkill(string skill)
    {
        var root = Path.Combine(SkillsFolder, skill);
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
                if (target.Length > 0 && !target.Contains("://", StringComparison.Ordinal) && target.EndsWith(".md", StringComparison.Ordinal))
                    pending.Push(Path.Combine(Path.GetDirectoryName(path)!, target));
            }
        }
        foreach (var reference in Directory.GetFiles(Path.Combine(root, "references"), "*.md", SearchOption.AllDirectories))
            Assert.Contains(Path.GetFullPath(reference), visited);
    }

    [Theory]
    [InlineData("excel-cli")]
    [InlineData("excel-mcp")]
    public void GeneratedGuidance_HasNoEmojiOrForcedUnrequestedWork(string skill)
    {
        foreach (var path in Directory.GetFiles(Path.Combine(SkillsFolder, skill), "*.md", SearchOption.AllDirectories))
        {
            var content = File.ReadAllText(path);
            Assert.False(content.EnumerateRunes().Any(rune =>
                rune.Value is >= 0x1F000 and <= 0x1FAFF or >= 0x2600 and <= 0x27BF or 0xFE0F),
                $"Emoji in {Path.GetRelativePath(SkillsFolder, path)}");
            foreach (var stale in new[] { "Always restore automatic mode", "Restore automatic mode",
                "NEVER Ask Clarifying Questions", "Write **one row at a time**", "MUST call `screenshot`",
                "pivottable(action: 'set-style')", "connection(action: 'test-connection')",
                "Query creation alone does not load", "imports the M code but does NOT execute it",
                "File name MUST match", "same STA thread pool", "retain a saved copy",
                "retain saved copies", "retain copies of both", "always work on copies",
                "before table operations to backup" })
            {
                Assert.DoesNotContain(stale, content, StringComparison.OrdinalIgnoreCase);
            }
        }
    }

    [Theory]
    [InlineData("excel-cli", "SKILL.md")]
    [InlineData("excel-mcp", @"references\calculation.md")]
    public void CalculationGuidance_RestoresPriorMode(string skill, string relativePath)
    {
        var content = File.ReadAllText(Path.Combine(SkillsFolder, skill, relativePath));
        Assert.Contains("get-mode", content);
        Assert.Contains("restore the prior mode", content, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("finally", content, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("excel-cli")]
    [InlineData("excel-mcp")]
    public void PermissionGuidance_IsSharedAndTaskScoped(string skill)
    {
        var root = Path.Combine(SkillsFolder, skill);
        var content = File.ReadAllText(Path.Combine(root, "references", "behavioral-rules.md"));
        Assert.Contains("clear, authorized request", content);
        Assert.Contains("ask one focused question", content);
        Assert.Contains("cleaning proposal is read-only", content);
        Assert.Contains("not user authorization", content);
        Assert.Contains("Do not create extra workbook copies or files as a safety step", content);
        Assert.Contains("temporary workbook objects", content);
        Assert.Contains("./references/behavioral-rules.md#intent-and-permission",
            File.ReadAllText(Path.Combine(root, "SKILL.md")));
        Assert.Equal(File.ReadAllText(Path.Combine(SkillsFolder, "shared", "behavioral-rules.md")).Replace("\r\n", "\n"),
            content.Replace("\r\n", "\n"));
    }

    [Theory]
    [InlineData("excel-cli")]
    [InlineData("excel-mcp")]
    public void VisibilityGuidance_ReusesPreferencesWithoutMandatoryQuestions(string skill)
    {
        var root = Path.Combine(SkillsFolder, skill);
        var policy = File.ReadAllText(Path.Combine(root, "references", "behavioral-rules.md"));
        Assert.Contains("known visibility preference", policy);
        Assert.Contains("Preserve an existing session's", policy);
        Assert.Contains("no known", policy);
        Assert.Contains("hidden by default", policy);
        Assert.Contains("Authentication may require visible Excel", policy);
        Assert.Contains("not show a", policy);
        var template = File.ReadAllText(Path.Combine(root, "SKILL.md"));
        Assert.Contains("known visibility preference", template);
        Assert.Contains("existing session's visibility", template.Replace("\r\n", "\n").Replace("\n  ", " "));
        Assert.Contains("does not mean showing a hidden Excel window", template);
    }

    [Fact]
    public void SharedGuidance_PreservesExactEntryPointNames()
    {
        (string Guide, string Mcp, string Cli)[] names =
        [
            ("analysis.md", "changing_cells", "--changing-cells"),
            ("drawing.md", "linked_cell", "--linked-cell"),
            ("excel_agent_mode.md", "save: true", "--save"),
            ("gotchas.md", "pivottable_field", "pivottablefield"),
            ("powerquery.md", "m_code_file", "--m-code-file"),
            ("querytable.md", "text_qualifier", "--text-qualifier"),
            ("range.md", "max_matches", "--max-matches"),
            ("screenshot.md", "sheet_name", "--sheet"),
            ("table.md", "has_headers: false", "--has-headers false"),
            ("workbook.md", "open_after_publish: false", "--open-after-publish false"),
            ("workflows.md", "datamodel_relationship", "datamodelrelationship"),
            ("xmlmap.md", "schema_file", "--schema-file")
        ];
        foreach (var skill in new[] { "excel-cli", "excel-mcp" })
        {
            foreach (var (guide, mcp, cli) in names)
            {
                var content = File.ReadAllText(Path.Combine(SkillsFolder, skill, "references", guide));
                Assert.Contains(mcp, content);
                Assert.Contains(cli, content);
            }
        }
        var cliReadme = File.ReadAllText(Path.Combine(SkillsFolder, "excel-cli", "references", "README.md"));
        Assert.DoesNotContain("translate them", cliReadme);
        Assert.Contains("native CLI examples", cliReadme);
    }

    [Theory]
    [InlineData("excel-cli")]
    [InlineData("excel-mcp")]
    public void ScreenshotGuidance_WorksheetIsOptional(string skill)
    {
        var content = File.ReadAllText(Path.Combine(SkillsFolder, skill, "references", "screenshot.md"));
        Assert.Equal(2, Regex.Matches(content, @"optional worksheet \(active sheet by default\)").Count);
    }

    [Theory]
    [InlineData("excel-cli")]
    [InlineData("excel-mcp")]
    public void QueryEvaluationGuidance_DistinguishesTemporaryChangesFromReads(string skill)
    {
        var content = File.ReadAllText(Path.Combine(SkillsFolder, skill, "references", "powerquery.md"));
        Assert.Contains("Evaluation is not a read-only operation", content);
        Assert.Contains("Execution may contact external sources", content);
        Assert.Contains("behavioral-rules.md#intent-and-permission", content);
    }

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
        foreach (var referencePath in CliCommandPages())
            AssertNoEmptyDescriptions(referencePath, "CLI command reference");
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CliCommandReference_HasCommands()
    {
        var content = ReadCliCommandPages();
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
        var content = ReadCliCommandPages();
        Assert.Contains("| Parameter | Description |", content);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void McpSkill_UsesDiscoveredSchemasAndTaskScopedGuidance()
    {
        var skillPath = Path.Combine(SkillsFolder, "excel-mcp", "SKILL.md");
        var content = File.ReadAllText(skillPath);
        Assert.Contains("Tool schemas describe the available actions and", content);
        Assert.Contains("Read-only tasks need no writes, formatting, Tables, charts, or PivotTables.", content);
        Assert.Contains("Do not convert every range automatically.", content);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CliCommandReference_HasActionsList()
    {
        var content = ReadCliCommandPages();
        Assert.Contains("**Actions:**", content);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CliCommandReference_CoversBranchAndGeneratedCommands()
    {
        var referenceContent = ReadCliCommandPages();
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
        var content = ReadCliCommandPages();

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
        var content = ReadCliCommandPages();
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
            .Append("index.md")
            .Append("README.md")
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        Assert.Equal(expectedFiles, fileNames);
        foreach (var sharedFile in expectedFiles.Except(["cli-commands.md", "index.md", "README.md"]))
        {
            var content = File.ReadAllText(Path.Combine(referencesPath, sharedFile));
            Assert.DoesNotContain("```mcp", content);
            Assert.DoesNotContain("```cli", content);
        }
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void BehavioralRules_CheckedInReferencesMatchSharedSource()
    {
        var shared = NormalizeLineEndings(File.ReadAllText(
            Path.Combine(SkillsFolder, "shared", "behavioral-rules.md")));
        var mcp = NormalizeLineEndings(File.ReadAllText(
            Path.Combine(SkillsFolder, "excel-mcp", "references", "behavioral-rules.md")));
        var cli = NormalizeLineEndings(File.ReadAllText(
            Path.Combine(SkillsFolder, "excel-cli", "references", "behavioral-rules.md")));

        Assert.Equal(shared, mcp);
        Assert.Equal(shared, cli);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void McpSkill_DoesNotDuplicateCalculationModeWorkflow()
    {
        var skillPath = Path.Combine(SkillsFolder, "excel-mcp", "SKILL.md");
        var content = File.ReadAllText(skillPath);

        Assert.Single(Regex.Matches(content, @"^## Bulk writes\r?$", RegexOptions.Multiline));
        Assert.Contains("calculation_mode(action: 'get-mode'", content);
        Assert.Contains("scope: 'workbook'", content);
        Assert.Contains("restore the prior mode", content);
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

    private static IEnumerable<Match> UnexpectedCamelCaseTokens(string content)
    {
        var allowedCamelCaseTokens = new HashSet<string>(StringComparer.Ordinal)
        {
            // MCP and CLI response properties.
            "canClose",
            "canOpen",
            "categoryRange",
            "chartName",
            "errorCategory",
            "errorMessage",
            "formulaPreview",
            "groupedFieldName",
            "isIrmProtected",
            "isPivotChart",
            "linkedPivotTable",
            "loadMode",
            "majorUnit",
            "matchingCells",
            "minorUnit",
            "newName",
            "oldName",
            "requiresVisibleSession",
            "returnedCount",
            "safeToCreate",
            "sessionId",
            "sourcePath",
            "sourceRange",
            "targetPath",
            "suggestedNextActions",
            "totalCount",
            "valuesRange",
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

        const string quotedText = @"'(?:\\.|[^'\\])*'|""(?:\\.|[^""\\])*""";
        var inputPositions = Regex.Matches(content, $@"\b[a-z][a-z_]*\((?:{quotedText}|[^)'""])*\)")
            .Cast<Match>()
            .SelectMany(call => Regex.Matches(call.Value,
                    $@"{quotedText}|(?<input>\b[a-z]+[A-Z][A-Za-z0-9]*)\s*:")
                .Cast<Match>()
                .Where(match => match.Groups["input"].Success)
                .Select(match => call.Index + match.Groups["input"].Index))
            .ToHashSet();

        return Regex.Matches(content, @"\b[a-z]+[A-Z][A-Za-z0-9]*\b")
            .Cast<Match>()
            .Where(match => !allowedCamelCaseTokens.Contains(match.Value) || inputPositions.Contains(match.Index));
    }

    [Theory]
    [InlineData("categoryRange")]
    [InlineData("sourceRange")]
    [InlineData("valuesRange")]
    [InlineData("sessionId")]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void McpSpellingGuard_RejectsResponsePropertyNamesAsInputs(string parameter)
    {
        var content = $"Response property `{parameter}`.\n```mcp\nchart_config(\n{parameter}: 'A1:A6')\n```";
        var unexpected = Assert.Single(UnexpectedCamelCaseTokens(content));
        Assert.Equal(parameter, unexpected.Value);
        Assert.True(unexpected.Index > content.IndexOf("chart_config", StringComparison.Ordinal));
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void McpSpellingGuard_AllowsResponseProseAndQuotedValues()
    {
        const string content = """
            Response properties: categoryRange, sourceRange, valuesRange, sessionId.
            ```mcp
            chart_config(action: 'set-title', title: 'categoryRange: (sourceRange)', session_id: sessionId)
            conditionalformat(action: 'add-rule', rule_type: 'colorScale')
            ```
            """;
        Assert.Empty(UnexpectedCamelCaseTokens(content));
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CanonicalMcpGuidance_UsesSnakeCaseInputs()
    {
        var canonicalFiles = Directory.GetFiles(Path.Combine(SkillsFolder, "excel-mcp", "references"), "*.md")
            .Append(Path.Combine(SkillsFolder, "templates", "SKILL.mcp.sbn"))
            .Append(Path.Combine(SkillsFolder, "excel-mcp", "references", "claude-desktop.md"));
        var unexpectedTokens = new List<string>();

        foreach (var path in canonicalFiles)
        {
            var relativePath = Path.GetRelativePath(SkillsFolder, path);
            var content = File.ReadAllText(path);
            foreach (var match in UnexpectedCamelCaseTokens(content))
            {
                var line = content[..match.Index].Count(character => character == '\n') + 1;
                unexpectedTokens.Add($"{relativePath}:{line}: {match.Value}");
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

    private static string NormalizeLineEndings(string content) =>
        content.Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n');

    private static string[] CliCommandPages() =>
        Directory.GetFiles(Path.Combine(SkillsFolder, "excel-cli", "references", "commands"), "*.md");

    private static string ReadCliCommandPages() =>
        string.Join('\n', CliCommandPages().Select(File.ReadAllText));

    [Fact]
    public void NativeCliExamples_UseOptionsFromLiveCommandHelp()
    {
        var root = Path.Combine(SkillsFolder, "excel-cli");
        foreach (var path in Directory.GetFiles(Path.Combine(root, "references"), "*.md"))
        {
            foreach (Match match in Regex.Matches(File.ReadAllText(path), @"(?m)^excelcli -q (?<command>[a-z]+) (?<action>[a-z-]+)(?<args>[^\r\n]*)"))
            {
                var command = match.Groups["command"].Value;
                var reference = File.ReadAllText(Path.Combine(root, "references", "commands", $"{command}.md"));
                Assert.Contains(match.Groups["action"].Value, reference);
                foreach (Match option in Regex.Matches(match.Groups["args"].Value, @"--[a-z][a-z0-9-]+"))
                    Assert.Contains($"`{option.Value}`", reference);
            }
        }
    }
}
