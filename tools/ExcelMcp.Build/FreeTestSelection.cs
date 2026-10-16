namespace Sbroenne.ExcelMcp.Build;

public sealed class FreeTestOptions
{
    public bool Local { get; set; }
    public bool HookTests { get; set; }
    public bool Contracts { get; set; }
    public bool SkillTests { get; set; }
    public bool PackagingTests { get; set; }
    public string[] ChangedPaths { get; set; } = [];
    public string? Group { get; set; }
    public string? PlanFile { get; set; }
    public string? ResultsDirectory { get; set; }
    public bool ListTests { get; set; }
}

public sealed record FreeTestProject(string Owner, string Filter);

public static class FreeTestSelection
{
    public static IReadOnlyList<FreeTestProject> Select(string root, FreeTestOptions options, ValidationPlan? savedPlan = null)
    {
        var selections = new Dictionary<string, string>(StringComparer.Ordinal);
        var onDemand = new Dictionary<string, string>(StringComparer.Ordinal);
        if (options.Group is { } group)
        {
            var commit = group == "Commit";
            if (commit) { group = "Fast"; }
            var plan = savedPlan ?? new ValidationPolicy(root).Select([], full: true);
            if (!plan.CiTestGroups.Contains(group, StringComparer.Ordinal)) { throw new InvalidOperationException($"The requested {group} group was not selected."); }
            var filters = ValidationExecution.Filters(plan, group);
            if (filters.Count == 0) { throw new InvalidOperationException($"{group} has no selected test projects."); }
            foreach (var (owner, filter) in filters)
            {
                if (commit && owner is not ("CLI" or "Core" or "McpServer" or "Service")) { continue; }
                if (string.IsNullOrWhiteSpace(filter)) { throw new InvalidOperationException($"Missing filter for {owner}."); }
                selections[owner] = group switch
                {
                    "Fast" => $"AdapterTestKind!=System&({filter})",
                    "Process" => $"AdapterTestKind=System&({filter})",
                    _ => filter
                };
                if (group == "Tooling" && plan.ToolingOnDemandFilters.TryGetValue(owner, out var selectedOnDemand))
                {
                    onDemand[owner] = selectedOnDemand;
                }
            }
        }
        else if (options.PlanFile is not null && !options.Local) { throw new ArgumentException("PlanFile requires an explicit Group or Local."); }
        else if (options.Local)
        {
            var plan = savedPlan ?? new ValidationPolicy(root).Select(options.ChangedPaths);
            foreach (var (owner, filter) in plan.FastFilters.Concat(plan.ProcessFilters)) { Add(owner, filter); }
            if (options.HookTests || plan.HookTests) { Add("ScriptSafety", plan.ToolingFilters.GetValueOrDefault("ScriptSafety") ?? "RequiresExcel=false"); }
            if (options.SkillTests || plan.SkillTests) { Add("SkillGeneration", plan.ToolingFilters.GetValueOrDefault("SkillGeneration") ?? "Feature=SkillGeneration"); }
            if (options.PackagingTests || plan.PackagingTests) { Add("Packaging", plan.ToolingFilters.GetValueOrDefault("Packaging") ?? "RequiresExcel=false"); }
            if (options.Contracts) { foreach (var owner in new[] { "Core", "CLI", "McpServer" }) { Add(owner, "Feature=GeneratedContracts"); } }
            foreach (var (owner, filter) in plan.ToolingOnDemandFilters) { onDemand[owner] = filter; }
        }
        else
        {
            foreach (var owner in new[] { "CLI", "ComInterop", "Core", "McpServer", "Service", "SkillGeneration", "Packaging", "ScriptSafety" }) { selections[owner] = "RequiresExcel=false"; }
        }
        return selections.Select(item => new FreeTestProject(item.Key, ExcelFreeFilter(item.Value, onDemand.GetValueOrDefault(item.Key)))).ToArray();

        void Add(string owner, string filter) => selections[owner] = selections.TryGetValue(owner, out var previous)
            ? string.Join('|', previous.Split('|').Concat(filter.Split('|')).Distinct(StringComparer.Ordinal)) : filter;
    }

    internal static string ExcelFreeFilter(string filter, string? onDemandFilter = null) =>
        onDemandFilter is null ? $"RequiresExcel=false&RunType!=OnDemand&({filter})"
        : $"RequiresExcel=false&((RunType!=OnDemand&({filter}))|(RunType=OnDemand&({onDemandFilter})))";

    public static async Task ExecuteAsync(string root, FreeTestOptions options, IProcessRunner runner, ValidationPlan? plan = null)
    {
        var selected = Select(root, options, plan);
        if (selected.Count == 0) { Console.Error.WriteLine("No local Excel-free tests selected."); return; }
        var results = options.ResultsDirectory ?? Path.Combine(root, "TestResults", $"excel-free-{Guid.NewGuid():N}");
        var execution = new TestExecution(root, runner);
        foreach (var item in selected) { await execution.RunAsync(item.Owner, item.Filter, results, excel: false, options.ListTests); }
    }
}
