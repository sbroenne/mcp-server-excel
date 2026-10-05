namespace Sbroenne.ExcelMcp.Build;

public sealed class ValidationPlan
{
    public int SchemaVersion { get; set; }
    public List<string> ChangedPaths { get; } = [];
    public bool Build { get; set; }
    public bool Excel { get; set; }
    public bool FullE2E { get; set; }
    public bool SourceChecks { get; set; }
    public bool DocumentationCounts { get; set; }
    public bool FullSolutionBuild { get; set; }
    public bool HookTests { get; set; }
    public bool SkillTests { get; set; }
    public bool PackagingTests { get; set; }
    public bool Cli { get; set; }
    public bool Mcp { get; set; }
    public bool Extension { get; set; }
    public bool ExtensionTests { get; set; }
    public bool Mcpb { get; set; }
    public bool Skills { get; set; }
    public bool Plugins { get; set; }
    public bool Docs { get; set; }
    public bool NpmTests { get; set; }
    public bool LockfileTests { get; set; }
    public bool InfrastructureDiagnostics { get; set; }
    public List<string> Reasons { get; } = [];
    public SortedDictionary<string, string> FastFilters { get; } = new(StringComparer.Ordinal);
    public SortedDictionary<string, string> ProcessFilters { get; } = new(StringComparer.Ordinal);
    public SortedDictionary<string, string> ToolingFilters { get; } = new(StringComparer.Ordinal);
    public List<ExcelSelection> ExcelSelections { get; } = [];
    public SortedSet<string> BuildProjects { get; } = new(StringComparer.Ordinal);
    public SortedSet<string> CodeQlLanguages { get; } = new(StringComparer.Ordinal);
    public SortedSet<string> ExcelGroups { get; } = new(StringComparer.Ordinal);
    public string[] FastProjects => [.. FastFilters.Keys];
    public string[] ProcessProjects => [.. ProcessFilters.Keys];
    public string[] ToolingProjects => [.. ToolingFilters.Keys];
    public string[] CiTestGroups => [
        .. FastFilters.Count > 0 ? new[] { "Fast" } : [],
        .. ProcessFilters.Count > 0 ? new[] { "Process" } : [],
        .. ToolingFilters.Count > 0 ? new[] { "Tooling" } : []
    ];
    public bool PackageBuild => Cli || Mcp || Extension;
    public bool Packages => PackageBuild || Mcpb || Skills || Plugins;
    public string SourceChecksGroup => SourceChecks && FastFilters.Count > 0 ? "Fast" : DocumentationCounts || SourceChecks ? "Tooling" : "";
}

public sealed record ExcelSelection(string Project, string Filter, string Area);
