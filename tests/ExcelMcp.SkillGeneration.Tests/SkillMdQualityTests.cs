using System.Diagnostics;
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
    public void McpSkill_UsesSchemasAndCanonicalInputNames()
    {
        var content = File.ReadAllText(Path.Combine(SkillsFolder, "excel-mcp", "SKILL.md"));
        Assert.Contains("server instructions already cover sessions", content, StringComparison.Ordinal);
        Assert.Contains("MCP inputs use the advertised snake_case names", content, StringComparison.Ordinal);
        var ranges = File.ReadAllText(Path.Combine(SkillsFolder, "excel-mcp", "references", "range.md"));
        Assert.Contains("session_id:", ranges, StringComparison.Ordinal);
        Assert.DoesNotContain("sessionId:", ranges, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("excel-cli")]
    [InlineData("excel-mcp")]
    public void Guidance_ExplainsCalculationOrderingAndDestructiveRecovery(string skill)
    {
        var content = NormalizeLineEndings(File.ReadAllText(
            Path.Combine(SkillsFolder, skill, "references", "behavioral-rules.md"))).Replace("\n", " ");
        Assert.Contains("concurrent requests", content, StringComparison.Ordinal);
        Assert.Contains("no guaranteed caller-defined order", content, StringComparison.Ordinal);
        Assert.Contains("manual needs explicit calculation", content, StringComparison.Ordinal);
        Assert.Contains("attempt to restore the prior mode", content, StringComparison.Ordinal);
        Assert.Contains("Restoration can fail without failing the write", content, StringComparison.Ordinal);
        Assert.Contains("what-if data tables, not ordinary worksheet Tables", content, StringComparison.Ordinal);
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
    public void CliReferences_UseTopicIndexWithoutDuplicatingNativeHelp()
    {
        var references = Path.Combine(SkillsFolder, "excel-cli", "references");
        var index = File.ReadAllText(Path.Combine(references, "index.md"));
        Assert.Contains("(range.md)", index);
        Assert.Contains("(report-formatting.md)", index);
        Assert.False(File.Exists(Path.Combine(references, "cli-commands.md")));
        Assert.False(Directory.Exists(Path.Combine(references, "commands")));
        Assert.Contains("--values", ReadCliHelp("range"));
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
        var windows = File.ReadAllText(Path.Combine(root, "references", "window.md"));
        Assert.Contains("preserve existing visibility", windows);
        Assert.Contains("does not request showing Excel", windows);
        if (skill == "excel-cli")
        {
            var template = File.ReadAllText(Path.Combine(root, "SKILL.md"));
            Assert.Contains("known visibility preference", template);
            Assert.Contains("existing session's visibility", template);
            Assert.Contains("does not mean showing a hidden Excel window", template);
        }
    }

    [Fact]
    public void SharedGuidance_PreservesExactEntryPointNames()
    {
        (string Guide, string Mcp, string Cli)[] names =
        [
            ("analysis.md", "changing_cells", "--changing-cells"),
            ("drawing.md", "linked_cell", "--linked-cell"),
            ("window.md", "save: true", "--save"),
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
            Assert.Contains("pivottable_field", File.ReadAllText(
                Path.Combine(SkillsFolder, "excel-mcp", "references", "pivottable.md")));
            Assert.Contains("pivottablefield", File.ReadAllText(
                Path.Combine(SkillsFolder, "excel-cli", "references", "pivottable.md")));
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
    public void NativeCliHelp_HasDescribedArgumentsAndOptions()
    {
        var commandGroups = NativeCliCommandGroups();
        Assert.True(commandGroups.Length >= 34, $"Expected all command groups, found {commandGroups.Length}.");
        foreach (var command in commandGroups)
        {
            var help = ReadCliHelp(command);
            Assert.Contains("USAGE:", help);
            var options = Regex.Matches(help, @"(?m)^ {4}(?:-[a-z], | {4})?--[a-z][a-z0-9-]*(?: <[^>]+>)?(?<description>[^\r\n]*)");
            Assert.NotEmpty(options);
            foreach (Match option in options)
                Assert.Matches(@" {2,}\S", option.Groups["description"].Value);
            if (command is not ("session" or "service" or "batch"))
            {
                Assert.Contains("<ACTION>", help);
                Assert.Contains("Available actions:", help);
                Assert.DoesNotContain("Available actions: OPTIONS:", Regex.Replace(help, @"\s+", " "));
            }
        }
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void McpSkill_LinksWorkflowDecisionsInsteadOfToolCatalogs()
    {
        var skillPath = Path.Combine(SkillsFolder, "excel-mcp", "SKILL.md");
        var content = File.ReadAllText(skillPath);
        Assert.Contains("./references/powerquery.md#recovering-a-failed-create", content);
        Assert.Contains("./references/range.md", content);
        Assert.Contains("./references/report-formatting.md", content);
        Assert.DoesNotContain("| Parameter | Description |", content);
        Assert.DoesNotContain("**Actions:**", content);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void McpSkill_UsesDiscoveredSchemasAndTaskScopedGuidance()
    {
        var skillPath = Path.Combine(SkillsFolder, "excel-mcp", "SKILL.md");
        var content = File.ReadAllText(skillPath);
        Assert.Contains("Tool schemas describe the available actions and", content);
        Assert.Contains("unrelated refresh, styling, or screenshots", content);
        var policy = File.ReadAllText(Path.Combine(SkillsFolder, "excel-mcp", "references", "behavioral-rules.md"));
        Assert.Contains("read-only unless the user requests", policy);
        Assert.Contains("not permission to create one", policy);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void NativeCliHelp_CoversLifecycleBranchesAndAliases()
    {
        Assert.Contains("open", ReadCliHelp("session"));
        Assert.Contains("stop", ReadCliHelp("service"));
        Assert.Contains("--show", ReadCliHelp("session", "open"));
        Assert.Contains("--save", ReadCliHelp("session", "close"));
        var rangeHelp = ReadCliHelp("range");
        Assert.Contains("--sheet", rangeHelp);
        Assert.Contains("--range", rangeHelp);
        Assert.DoesNotContain("--sheet-name", rangeHelp);
        Assert.DoesNotContain("--range-address", rangeHelp);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CliSkill_DelegatesFullCommandReference()
    {
        var skillPath = Path.Combine(SkillsFolder, "excel-cli", "SKILL.md");
        var content = File.ReadAllText(skillPath);

        Assert.Contains("excelcli --help", content);
        Assert.Contains("excelcli <command> --help", content);
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
        Assert.Contains("./references/report-formatting.md", content);
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
            .Append("index.md")
            .Append("README.md")
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        Assert.Equal(expectedFiles, fileNames);
        foreach (var sharedFile in expectedFiles.Except(["index.md", "README.md"]))
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

        Assert.DoesNotContain("calculation_mode(action:", content);
        var calculation = File.ReadAllText(Path.Combine(SkillsFolder, "excel-mcp", "references", "calculation.md"));
        Assert.Contains("calculation_mode(action: 'get-mode'", calculation);
        Assert.Contains("scope: 'workbook'", calculation);
        Assert.Contains("restore the prior mode", calculation, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("one rectangular write is already batched", calculation);
    }

    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    [Theory]
    [InlineData("excel-cli")]
    [InlineData("excel-mcp")]
    public void ReportFormatting_PreservesOptionalFinancialConventionsAndDashboardAdvice(string skill)
    {
        var content = File.ReadAllText(Path.Combine(SkillsFolder, skill, "references", "report-formatting.md"));
        Assert.Contains("not a requirement for reads, raw exports, or targeted data", content);
        Assert.Contains("## Optional financial-model conventions", content);
        foreach (var colour in new[] { "#0000FF", "#000000", "#008000", "#FF0000", "#FFFF00" })
            Assert.Contains(colour, content);
        Assert.Contains("Record the source, date, and specific reference", content);
        Assert.Contains("## Dashboard layout", content);
        Assert.Contains("not a mandatory screenshot step", content);
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

    private static readonly Dictionary<string, string> CliHelpCache = new(StringComparer.Ordinal);

    private static string ReadCliHelp(params string[] command)
    {
        var key = string.Join(' ', command);
        if (CliHelpCache.TryGetValue(key, out var cached))
            return cached;
        var executable = Path.Combine(GeneratedAssetsFixture.RepositoryDirectory,
            "src", "ExcelMcp.CLI", "bin", "Release", "net10.0-windows", "excelcli.dll");
        Assert.True(File.Exists(executable), "Build the Release CLI before testing native help.");
        var info = new ProcessStartInfo("dotnet")
        {
            WorkingDirectory = GeneratedAssetsFixture.RepositoryDirectory,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        foreach (var argument in new[] { executable, "-q" }.Concat(command).Append("--help"))
            info.ArgumentList.Add(argument);
        using var process = Process.Start(info)!;
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(30000))
        {
            process.Kill(entireProcessTree: true);
            process.WaitForExit();
            throw new TimeoutException($"CLI help for '{key}' exceeded 30 seconds.");
        }
        var content = stdout.GetAwaiter().GetResult();
        var error = stderr.GetAwaiter().GetResult();
        Assert.True(process.ExitCode == 0, $"CLI help for '{key}' failed:\n{content}\n{error}");
        Assert.Equal(string.Empty, error);
        CliHelpCache.Add(key, content);
        return content;
    }

    private static string[] NativeCliCommandGroups() =>
        Regex.Matches(ReadCliHelp().Split("COMMANDS:", StringSplitOptions.None)[1],
                @"(?m)^ {4}([a-z][a-z0-9-]+)(?: <ACTION>)? {2,}\S")
            .Select(match => match.Groups[1].Value)
            .Distinct(StringComparer.Ordinal)
            .ToArray();

    private static string NormalizeLineEndings(string content) =>
        content.Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n');

    [Fact]
    public void NativeCliExamples_UseOptionsFromLiveCommandHelp()
    {
        var root = Path.Combine(SkillsFolder, "excel-cli");
        var paths = Directory.GetFiles(Path.Combine(root, "references"), "*.md")
            .Append(Path.Combine(root, "SKILL.md"));
        var exampleCount = 0;
        foreach (var path in paths)
        {
            foreach (Match match in Regex.Matches(File.ReadAllText(path), @"(?m)^excelcli -q (?<command>[a-z]+) (?<action>[a-z-]+)(?<args>[^\r\n]*)"))
            {
                var command = match.Groups["command"].Value;
                exampleCount++;
                var reference = ReadCliHelp(command);
                if (command is "session" or "service")
                    reference = ReadCliHelp(command, match.Groups["action"].Value);
                else
                {
                    var arguments = Regex.Replace(reference, @"\s+", " ")
                        .Split("OPTIONS:", StringSplitOptions.None)[0];
                    Assert.Matches($@"(?<![\w-]){Regex.Escape(match.Groups["action"].Value)}(?![\w-])", arguments);
                }
                foreach (Match option in Regex.Matches(match.Groups["args"].Value, @"--[a-z][a-z0-9-]+"))
                    Assert.Matches($@"(?<![\w-]){Regex.Escape(option.Value)}(?![\w-])", reference);
            }
        }
        Assert.True(exampleCount >= 30, $"Expected native examples across the guides, found {exampleCount}.");
    }
}
