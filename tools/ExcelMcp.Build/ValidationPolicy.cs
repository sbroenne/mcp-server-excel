using System.Text.RegularExpressions;

namespace Sbroenne.ExcelMcp.Build;

public sealed partial class ValidationPolicy(string root)
{
    private TestCatalog? _catalog;
    private TestCatalog Catalog => _catalog ??= new TestCatalog(root);
    private static readonly string[] RuntimeOwners = ["CLI", "ComInterop", "Core", "McpServer", "Service"];
    private static readonly string[] ToolingOwners = ["Packaging", "ScriptSafety", "SkillGeneration"];

    public ValidationPlan Select(IEnumerable<string> paths, bool full = false)
    {
        ArgumentNullException.ThrowIfNull(paths);
        var plan = new ValidationPlan { SchemaVersion = 1 };
        if (full)
        {
            plan.Reasons.Add("Complete validation explicitly selected.");
            plan.FullE2E = true;
            foreach (var owner in RuntimeOwners.Concat(ToolingOwners))
            {
                AddClasses(plan, Catalog.ForOwner(owner), "All");
            }
            plan.FullE2E = true;
            plan.FullSolutionBuild = true;
            plan.SourceChecks = true;
            plan.DocumentationCounts = true;
            plan.Docs = true;
            plan.ExtensionTests = true;
            plan.Cli = plan.Mcp = plan.Extension = plan.Mcpb = plan.Skills = plan.Plugins = true;
            plan.NpmTests = plan.LockfileTests = plan.InfrastructureDiagnostics = true;
            foreach (var language in new[] { "actions", "csharp", "javascript-typescript", "python" })
            {
                plan.CodeQlLanguages.Add(language);
            }
            FinalizePlan(plan);
            return plan;
        }
        foreach (var path in paths.Select(path => path?.Replace('\\', '/') ?? throw new ArgumentException("Validation paths cannot be null."))
            .Distinct(StringComparer.Ordinal).Order(StringComparer.Ordinal))
        {
            if (string.IsNullOrWhiteSpace(path) || Path.IsPathRooted(path) || path.Split('/').Any(part => part is "." or ".."))
            {
                throw new ArgumentException($"Validation inputs must be repository-relative paths: {path}.");
            }
            AddLanguage(plan, path);
            plan.ChangedPaths.Add(path);
            SelectPath(plan, path);
        }
        FinalizePlan(plan);
        return plan;
    }

    private void SelectPath(ValidationPlan plan, string path)
    {
        var reasonCount = plan.Reasons.Count;
        var feature = CommandPath().Match(path);
        if (SelectDocumentation(plan, path))
        {
            plan.Reasons.Add($"{path} -> documentation and its distribution owners");
        }
        else if (feature.Success)
        {
            var area = feature.Groups[1].Value;
            var features = Features(area);
            var contract = Path.GetFileName(path).StartsWith('I') ||
                path.Contains("/Attributes/", StringComparison.Ordinal);
            AddFeature(plan, "Core", features, area);
            AddFeature(plan, "Service", features, area);
            if (area == "Diag") { AddFeature(plan, "CLI", ["Diag"], area); }
            if (area != "Diag" && !Catalog.ForFeatures("Core", features).Concat(Catalog.ForFeatures("Service", features)).Any())
            {
                throw new InvalidOperationException($"No validation mapping for command area {area} ({path}).");
            }
            if (contract)
            {
                AddFeature(plan, "Core", ["GeneratedContracts"], "GeneratedContracts");
                AddFeature(plan, "CLI", ["GeneratedContracts"], "GeneratedContracts");
                AddFeature(plan, "McpServer", ["GeneratedContracts"], "GeneratedContracts");
                AddClasses(plan, Catalog.ForFeatures("CLI", features).Where(type => type.RequiredOnly && type.Name.StartsWith("CliNative", StringComparison.Ordinal)), $"{area} native parameter boundary");
                plan.DocumentationCounts = true;
                plan.FullSolutionBuild = true;
            }
            if (area == "DataModel")
            {
                AddFeature(plan, "Service", ["PowerQuery", "PivotTables", "Slicer", "Tables"], "DataModel dependencies");
            }
            if (area is "Connection" or "QueryTable")
            {
                AddFeature(plan, "Service", ["PowerQuery"], "PowerQuery refresh dependency");
            }
            plan.SourceChecks = true;
            plan.Cli = plan.Mcp = plan.Extension = true;
            plan.BuildProjects.Add("src/ExcelMcp.Core/ExcelMcp.Core.csproj");
            plan.Reasons.Add($"{path} -> {area}{(contract ? " and generated consumers" : " implementation")}");
        }
        else if (TestProjectPath().Match(path) is { Success: true } test)
        {
            var owner = test.Groups[1].Value;
            if (path.EndsWith(".cs", StringComparison.Ordinal))
            {
                var classes = Catalog.ForFile(path).ToArray();
                if (classes.Length == 0)
                {
                    throw new InvalidOperationException($"No validation mapping for test source {path}; identify its consuming tests.");
                }
                AddClasses(plan, classes, owner);
            }
            else
            {
                AddClasses(plan, Catalog.ForOwner(owner), owner);
            }
            plan.BuildProjects.Add($"tests/ExcelMcp.{owner}.Tests/ExcelMcp.{owner}.Tests.csproj");
            plan.Reasons.Add($"{path} -> owning {owner} cases");
        }
        else if (Matches(path, @"^tests/Directory\."))
        {
            foreach (var owner in RuntimeOwners.Concat(ToolingOwners).Append("Diagnostics")) { AddClasses(plan, Catalog.ForOwner(owner), "Test build inputs"); }
            plan.FullSolutionBuild = true;
            plan.Reasons.Add($"{path} -> all test-project settings consumers");
        }
        else if (path.StartsWith("tests/Shared/", StringComparison.Ordinal))
        {
            var classes = Catalog.ForFile(path).ToArray();
            if (classes.Length == 0) { throw new InvalidOperationException($"No validation mapping for shared fixture {path}."); }
            AddClasses(plan, classes, "Shared fixture consumers");
            plan.Reasons.Add($"{path} -> shared fixture consumers");
        }
        else if (Matches(path, @"^tools/ExcelMcp\.Build/(PackageFiles|AuthoredPackages|RuntimePackages|PackageExecution)\.cs$"))
        {
            AddFeature(plan, "Packaging", ["Packaging", "McpbPackaging", "PluginBootstrap", "PluginSkillVersion", "PluginPublication"], "Typed package operations");
            AddFeature(plan, "SkillGeneration", ["SkillGeneration"], "Typed prepared skills");
            AddClasses(plan, Catalog.ForOwner("ScriptSafety").Where(type => type.Name == "TypedPackageMigrationTests"), "Package adapters");
            plan.Cli = plan.Mcp = plan.Extension = plan.Mcpb = plan.Skills = plan.Plugins = true;
            plan.Reasons.Add($"{path} -> direct package and prepared-skill consumers, not workbook tests");
        }
        else if (Matches(path, @"^(tools/ExcelMcp\.Build/|build\.ps1$|scripts/(Get-ValidationPlan|Get-CiValidationPlan|Build-CiInputs|Invoke-BuildTool|Invoke-ExcelFreeTests|Invoke-ExcelTests|Get-ExcelTestGroups|Invoke-TestStage|Test-CiCompletion|pre-commit)\.ps1$|\.github/workflows/ci\.yml$)"))
        {
            AddFeature(plan, "ScriptSafety", ["PreCommit"], "Build and selection");
            plan.Reasons.Add($"{path} -> build/selection regressions, not workbook features");
        }
        else if (path == ".editorconfig")
        {
            foreach (var owner in RuntimeOwners.Concat(ToolingOwners).Append("Diagnostics"))
            {
                AddClasses(plan, Catalog.ForOwner(owner).Where(type => type.Name == "TestClassificationArchitectureTests"), "Build analysis");
            }
            plan.FullSolutionBuild = plan.SourceChecks = true;
            plan.Reasons.Add($"{path} -> build analysis and test classification");
        }
        else if (path.StartsWith("src/ExcelMcp.Core/", StringComparison.Ordinal) && path.EndsWith(".cs", StringComparison.Ordinal) &&
            !Matches(path, @"^src/ExcelMcp\.Core/(Attributes/|Models/Actions/|AssemblyAttributes\.cs$|Commands/(I?FileCommands|CoreLookupHelpers)\.cs$)"))
        {
            var consumers = new SourceDependencies(root, Catalog).CommandConsumers(path);
            foreach (var area in consumers.Areas)
            {
                AddFeature(plan, "Core", Features(area), $"{area} shared-source dependency");
                AddFeature(plan, "Service", Features(area), $"{area} shared-source dependency");
            }
            AddClasses(plan, Catalog.ForReferences("Core", consumers.Names), "Shared-source unit consumers");
            if (consumers.Areas.Length == 0 && !Catalog.ForReferences("Core", consumers.Names).Any())
            {
                throw new InvalidOperationException($"No validation mapping for shared Core source {path}; add its actual consumers.");
            }
            plan.SourceChecks = true;
            plan.Cli = plan.Mcp = plan.Extension = true;
            plan.BuildProjects.Add("src/ExcelMcp.Core/ExcelMcp.Core.csproj");
            plan.Reasons.Add($"{path} -> dependent command areas: {string.Join(", ", consumers.Areas)}");
        }
        else if (Matches(path, @"^(Directory\.|global\.json$|NuGet\.Config$|Sbroenne\.ExcelMcp\.sln$|tests/Directory\.|src/ExcelMcp\.(Core|ComInterop|Service|Cleanup)/)"))
        {
            var owners = path.StartsWith("src/", StringComparison.Ordinal)
                ? RuntimeOwners
                : RuntimeOwners.Concat(ToolingOwners).Append("Diagnostics");
            plan.FullE2E = true;
            foreach (var owner in owners) { AddClasses(plan, Catalog.ForOwner(owner), "Shared runtime/build"); }
            plan.SourceChecks = true;
            plan.FullE2E = true;
            plan.InfrastructureDiagnostics = true;
            plan.FullSolutionBuild = true;
            plan.Cli = plan.Mcp = plan.Extension = true;
            plan.Reasons.Add($"{path} -> shared runtime/build consumers");
        }
        else if (path.StartsWith("src/ExcelMcp.Generators", StringComparison.Ordinal))
        {
            foreach (var owner in new[] { "Core", "CLI", "McpServer" })
            {
                AddFeature(plan, owner, ["GeneratedContracts"], "GeneratedContracts");
            }
            plan.SourceChecks = plan.DocumentationCounts = plan.FullSolutionBuild = true;
            plan.Cli = plan.Mcp = plan.Extension = true;
            plan.Reasons.Add($"{path} -> generated contract consumers");
        }
        else if (path.StartsWith("src/ExcelMcp.Diagnostics/", StringComparison.Ordinal))
        {
            AddClasses(plan, Catalog.ForOwner("Diagnostics"), "Diagnostics");
            plan.BuildProjects.Add("src/ExcelMcp.Diagnostics/ExcelMcp.Diagnostics.csproj");
            plan.Reasons.Add($"{path} -> diagnostics consumers");
        }
        else if (Matches(path, @"^src/ExcelMcp\.(CLI|McpServer)/") && !path.EndsWith(".md", StringComparison.Ordinal))
        {
            var owner = path.StartsWith("src/ExcelMcp.CLI/", StringComparison.Ordinal) ? "CLI" : "McpServer";
            AddClasses(plan, new SourceDependencies(root, Catalog).AdapterConsumers(owner, path), $"{owner} adapter consumers");
            if (owner == "CLI") { plan.Cli = true; }
            else { plan.Mcp = plan.Extension = true; }
            plan.SourceChecks = true;
            plan.Reasons.Add($"{path} -> dependent {owner} adapter cases, not the entire adapter project");
        }
        else if (Matches(path, @"^(skills/|docs/reference/report-formatting\.md$|scripts/Build-AgentSkills\.ps1$|docs/AGENT-SKILLS\.md$)"))
        {
            AddFeature(plan, "SkillGeneration", ["SkillGeneration"], "Skill preparation");
            plan.Skills = true;
            if (path != "docs/AGENT-SKILLS.md") { plan.Plugins = plan.Extension = true; }
            plan.Reasons.Add($"{path} -> prepared skill consumers");
        }
        else if (Matches(path, @"^(\.github/plugins/|scripts/(Build-Plugins|Sync-PublishedPluginRepo|Publish-PreparedPlugins)\.ps1$|scripts/(PluginContent|AwesomeCopilotPolicy|Update-AwesomeCopilot)\.mjs$|\.github/workflows/(publish-plugins\.yml|update-awesome-copilot\.(md|lock\.yml))$)"))
        {
            AddFeature(plan, "Packaging", ["PluginPublication", "PluginBootstrap", "PluginSkillVersion"], "Plugin packaging");
            plan.Plugins = true;
            plan.Reasons.Add($"{path} -> plugin packaging, not runtime tests");
        }
        else if (Matches(path, @"^(doc-counts\.json$|scripts/check-doc-counts\.ps1$)"))
        {
            plan.ToolingFilters["Packaging"] = plan.ToolingFilters.TryGetValue("Packaging", out var previous)
                ? Union(previous, "FullyQualifiedName~DocumentationCounts") : "FullyQualifiedName~DocumentationCounts";
            plan.DocumentationCounts = plan.FullSolutionBuild = true;
            plan.Reasons.Add($"{path} -> documentation count derivation");
        }
        else if (Matches(path, @"^mcpb/"))
        {
            AddFeature(plan, "Packaging", ["McpbPackaging"], "Mcpb packaging");
            plan.Mcpb = true;
            plan.Reasons.Add($"{path} -> metadata bundle");
        }
        else if (Matches(path, @"^(scripts/(Build-NpmPackages|Test-NpmPackages|Build-ReleasePackages|PackageHelpers|Build-Changelog|Update-(ReleaseVersion|McpRegistry)Metadata|Resolve-McpRegistryRelease|Test-McpRegistryPublication)\.ps1$|\.github/workflows/(release|publish-mcp-registry)\.yml$|package(-lock)?\.json$|\.npmrc$)"))
        {
            AddFeature(plan, "Packaging", ["Packaging", "McpbPackaging", "ReleaseMetadata", "PluginSkillVersion"], "Package preparation");
            plan.Cli = plan.Mcp = plan.Extension = plan.Mcpb = plan.Skills = plan.Plugins = true;
            plan.Reasons.Add($"{path} -> package preparation consumers");
        }
        else if (Matches(path, @"^scripts/check-(com-leaks|success-flag|dynamic-casts|workbook-package-access)\.ps1$"))
        {
            AddClasses(plan, Catalog.ForOwner("ScriptSafety").Where(type => type.Name is "AutomationSafetyTests" or "WorkbookPackageAccessGuardTests" or "TypedSourceGuardTests"), "Source safeguards");
            plan.SourceChecks = true;
            plan.Reasons.Add($"{path} -> owning source safeguards");
        }
        else if (Matches(path, @"^scripts/(Invoke-CopilotSetupNpm|Install-CopilotPonytailReview)\.ps1$|^\.github/workflows/copilot-setup-steps\.yml$|^scripts/(AzureRunnerHost|ExcelRunner(Host|Policy)|Invoke-ExcelRunner(Control|Maintenance))\.ps1$|^scripts/(Deploy|Initialize|Install|Open|Register)-ExcelAgent(Runner|Desktop|Office|Toolchain|Activation)\.ps1$|^infrastructure/azure/[^/]*excel[^/]*\.(ps1|bicep)$|^\.github/workflows/excel-runner[^/]*\.yml$"))
        {
            plan.Reasons.Add($"{path} -> operational maintenance, outside automated validation");
        }
        else if (Matches(path, @"^(scripts/(check-|Test-NpmLockfiles)|scripts/tests/|infrastructure/azure/(configure-analytics-oidc|deploy-appinsights)\.ps1$|videos/excel-mcp-intro/Capture-Evidence\.ps1$)"))
        {
            AddFeature(plan, "ScriptSafety", ["AutomationSafety"], "Owning script safety");
            plan.Reasons.Add($"{path} -> script safety only");
        }
        else if (Matches(path, @"^scripts/(Test-E2E|Test-CliWorkflow|Test-CliApiCoverage|Stop-ExcelMcpProcesses)\.ps1$"))
        {
            plan.FullE2E = true;
            AddClasses(plan, Catalog.ForOwner("CLI").Where(type => type.Name is "CliWorkflowAcceptanceTests" or "PreBuildGracefulSaveAcceptanceTests"), "Acceptance");
            AddClasses(plan, Catalog.ForOwner("McpServer").Where(type => type.Name == "McpServerSmokeTests"), "Acceptance");
            plan.Reasons.Add($"{path} -> affected acceptance/cleanup boundary");
        }
        else if (path.StartsWith("npm-packages/", StringComparison.Ordinal))
        {
            plan.NpmTests = true;
            plan.Cli |= path.Contains("excelcli", StringComparison.Ordinal) || path.Contains("/shared/", StringComparison.Ordinal);
            plan.Mcp |= path.Contains("mcp-server-excel", StringComparison.Ordinal) || path.Contains("/shared/", StringComparison.Ordinal);
            plan.Reasons.Add($"{path} -> npm launcher and affected packages");
        }
        else if (path.StartsWith("vscode-extension/", StringComparison.Ordinal))
        {
            plan.Extension = true;
            plan.ExtensionTests = true;
            plan.Reasons.Add($"{path} -> extension checks and package");
        }
        else if (path.StartsWith("samples/world-bank-dashboard/", StringComparison.Ordinal))
        {
            plan.Reasons.Add($"{path} -> published sample and documentation checks");
        }
        else if (path.EndsWith(".md", StringComparison.Ordinal) || Matches(path, @"^(docs/|gh-pages/|\.github/|\.changeset/|infrastructure/|videos/|llm-tests/|scripts/.*(UsageAnalytics|StarHistory)|\.(gitignore|gitattributes)$|LICENSE$)"))
        {
            plan.Reasons.Add($"{path} -> non-runtime input");
        }
        if (plan.Reasons.Count == reasonCount)
        {
            throw new InvalidOperationException($"No validation mapping for {path}. Add an owning area; a full-suite fallback is not allowed.");
        }
        if (Matches(path, @"^(docs/(?!agents/)|gh-pages/|samples/world-bank-dashboard/|README\.md$|FEATURES\.md$|CHANGELOG\.md$|LICENSE$|doc-counts\.json$)"))
        {
            plan.Docs = true;
        }
        if (Matches(path, @"(^|/)(package-lock\.json|\.npmrc)$|^scripts/(Test-NpmLockfiles|check-npm-lockfiles)\.ps1$"))
        {
            plan.LockfileTests = true;
        }
    }

    private void AddFeature(ValidationPlan plan, string owner, string[] features, string area)
    {
        var classes = Catalog.ForFeatures(owner, features).ToArray();
        AddClasses(plan, classes, area);
        if (classes.Any(type => type.Excel || type.ExcelFree))
        {
            plan.Reasons.Add($"{owner}: {area} behavior and boundary cases.");
        }
    }

    private static void AddClasses(ValidationPlan plan, IEnumerable<TestClass> classes, string area)
    {
        foreach (var group in classes.GroupBy(type => type.Owner))
        {
            var owner = group.Key;
            var free = group.Where(type => type.ExcelFree && !type.System).Select(type => type.FullName).ToArray();
            var system = group.Where(type => type.ExcelFree && type.System).Select(type => type.FullName).ToArray();
            var excel = group.Where(type => type.Excel && (!plan.FullE2E || !type.RequiredOnly)).Select(type => type.FullName).ToArray();
            AddNames(ToolingOwners.Contains(owner, StringComparer.Ordinal) ? plan.ToolingFilters : plan.FastFilters, owner, free);
            AddNames(plan.ProcessFilters, owner, system);
            if (excel.Length > 0)
            {
                var filter = ClassFilter(excel);
                var existing = plan.ExcelSelections.FindIndex(selection => selection.Project == owner && selection.Area == area);
                if (existing < 0) { plan.ExcelSelections.Add(new ExcelSelection(owner, filter, area)); }
                else
                {
                    var previous = plan.ExcelSelections[existing];
                    plan.ExcelSelections[existing] = previous with { Filter = Union(previous.Filter, filter) };
                }
                plan.ExcelGroups.Add(area);
            }
            if (free.Length + system.Length + excel.Length > 0 ||
                plan.FullE2E && group.Any(type => type.Excel && type.RequiredOnly))
            {
                plan.BuildProjects.Add($"tests/ExcelMcp.{owner}.Tests/ExcelMcp.{owner}.Tests.csproj");
            }
        }
    }

    private static void AddNames(SortedDictionary<string, string> filters, string owner, IEnumerable<string> names)
    {
        var values = names.Distinct(StringComparer.Ordinal).Order(StringComparer.Ordinal).ToArray();
        if (values.Length == 0) { return; }
        var filter = ClassFilter(values);
        filters[owner] = filters.TryGetValue(owner, out var previous) ? Union(previous, filter) : filter;
    }

    private static string ClassFilter(IEnumerable<string> names) =>
        string.Join('|', names.Select(name => $"FullyQualifiedName~{name}."));

    private static string Union(string first, string second) =>
        string.Join('|', first.Split('|').Concat(second.Split('|')).Distinct(StringComparer.Ordinal).Order(StringComparer.Ordinal));

    private void FinalizePlan(ValidationPlan plan)
    {
        if (plan.SourceChecks && plan.FastFilters.Count == 0 && plan.ToolingFilters.Count == 0)
        {
            AddClasses(plan, Catalog.ForOwner("ScriptSafety").Where(type => type.Name == "WorkbookPackageAccessGuardTests"), "Source safeguards");
        }
        plan.Excel = plan.ExcelSelections.Count > 0 || plan.FullE2E;
        plan.HookTests = plan.ToolingFilters.ContainsKey("ScriptSafety");
        plan.SkillTests = plan.ToolingFilters.ContainsKey("SkillGeneration");
        plan.PackagingTests = plan.ToolingFilters.ContainsKey("Packaging");
        plan.Build = plan.BuildProjects.Count > 0 || plan.FullSolutionBuild;
    }

    private static void AddLanguage(ValidationPlan plan, string path)
    {
        if (Matches(path, @"\.(cs|csproj|sln|slnf|props|targets)$|(^|/)(global\.json|NuGet\.Config)$")) { plan.CodeQlLanguages.Add("csharp"); }
        if (Matches(path, @"\.(js|jsx|mjs|cjs|ts|tsx|mts|cts)$|(^|/)(package(-lock)?\.json|[jt]sconfig[^/]*\.json|\.npmrc)$")) { plan.CodeQlLanguages.Add("javascript-typescript"); }
        if (Matches(path, @"\.py$|(^|/)(requirements[^/]*\.txt|pyproject\.toml|poetry\.lock|uv\.lock|Pipfile(\.lock)?|setup\.cfg)$")) { plan.CodeQlLanguages.Add("python"); }
        if (Matches(path, @"^\.github/(workflows/[^/]+\.ya?ml$|actions/.+/action\.ya?ml$)")) { plan.CodeQlLanguages.Add("actions"); }
        if (Matches(path, @"^\.github/(workflows/codeql\.yml|codeql/)|^(tools/ExcelMcp\.Build/|scripts/Get-(Ci)?ValidationPlan\.ps1$)"))
        {
            foreach (var language in new[] { "actions", "csharp", "javascript-typescript", "python" }) { plan.CodeQlLanguages.Add(language); }
        }

    }

    public static string[] LanguagesForPath(string path)
    {
        var plan = new ValidationPlan();
        AddLanguage(plan, path.Replace('\\', '/'));
        return [.. plan.CodeQlLanguages];
    }

    private static string[] Features(string area) => area switch
    {
        "Range" => ["Range", "Ranges", "FineFormatting", "Protection", "NativeDataCleanup"],
        "Sheet" => ["Sheet", "Worksheets", "StructuredFilters", "PageLayout", "Protection"],
        "Table" => ["Table", "Tables", "TableStyles"],
        "Filtering" => ["StructuredFilters", "Range", "Tables", "Worksheets"],
        "NamedRange" => ["Parameters", "NamedRange"],
        "Chart" => ["Chart", "Charts", "ChartDepth"],
        "PivotTable" => ["PivotTable", "PivotTables", "PivotDepth", "PivotCalculation"],
        "Slicer" => ["Slicer", "Timelines"],
        "Drawing" => ["Drawing", "DrawingLayout"],
        "ConditionalFormat" => ["ConditionalFormat", "ConditionalRuleEditing"],
        "Calculation" => ["CalculationMode"],
        "Workbook" => ["Workbook", "WorkbookTheme", "CellStyles", "PageLayout"],
        "File" => ["Files", "Session", "SessionLifecycle"],
        "Vba" => ["VBA", "VBATrust"],
        _ => [area]
    };

    private static bool Matches(string value, string pattern) => Regex.IsMatch(value, pattern, RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
    private bool SelectDocumentation(ValidationPlan plan, string path)
    {
        if (!path.EndsWith(".md", StringComparison.Ordinal) && path != "LICENSE") { return false; }
        if (Matches(path, @"^(skills/|\.github/plugins/|docs/AGENT-SKILLS\.md$|docs/reference/report-formatting\.md$|\.github/workflows/update-awesome-copilot\.md$)")) { return false; }
        if (Matches(path, @"(^|/)(AGENTS|CLAUDE)\.md$|^\.github/copilot-instructions\.md$")) { return true; }
        var cli = Matches(path, @"^README\.md$|^src/ExcelMcp\.CLI/README\.md$|^npm-packages/excelcli[^/]*/README\.md$|^(CHANGELOG\.md|LICENSE)$");
        var mcp = Matches(path, @"^README\.md$|^src/ExcelMcp\.McpServer/README\.md$|^npm-packages/mcp-server-excel[^/]*/README\.md$|^(CHANGELOG\.md|LICENSE)$");
        var extension = Matches(path, @"^vscode-extension/(README\.md|LICENSE|CHANGELOG\.md)$|^(CHANGELOG\.md|LICENSE)$");
        var mcpb = path is "LICENSE" or "CHANGELOG.md";
        plan.Cli |= cli;
        plan.Mcp |= mcp;
        plan.Extension |= extension;
        plan.Mcpb |= mcpb;
        if (cli || mcp || extension || mcpb)
        {
            AddFeature(plan, "Packaging", ["Packaging"], "Distributed documentation");
        }
        return true;
    }
    [GeneratedRegex(@"^src/ExcelMcp\.Core/Commands/([^/]+)/", RegexOptions.CultureInvariant)]
    private static partial Regex CommandPath();
    [GeneratedRegex(@"^tests/ExcelMcp\.([^/]+)\.Tests/", RegexOptions.CultureInvariant)]
    private static partial Regex TestProjectPath();
}
