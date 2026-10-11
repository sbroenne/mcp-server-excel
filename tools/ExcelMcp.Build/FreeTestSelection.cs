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
        if (options.Group is { } group)
        {
            var plan = savedPlan ?? new ValidationPolicy(root).Select([], full: true);
            if (!plan.CiTestGroups.Contains(group, StringComparer.Ordinal)) { throw new InvalidOperationException($"The requested {group} group was not selected."); }
            var filters = ValidationExecution.Filters(plan, group);
            if (filters.Count == 0) { throw new InvalidOperationException($"{group} has no selected test projects."); }
            foreach (var (owner, filter) in filters)
            {
                if (string.IsNullOrWhiteSpace(filter)) { throw new InvalidOperationException($"Missing filter for {owner}."); }
                selections[owner] = group switch
                {
                    "Fast" => $"AdapterTestKind!=System&({filter})",
                    "Process" => $"AdapterTestKind=System&({filter})",
                    _ => filter
                };
            }
        }
        else if (options.PlanFile is not null) { throw new ArgumentException("PlanFile requires an explicit Group."); }
        else if (options.Local)
        {
            var plan = new ValidationPolicy(root).Select(options.ChangedPaths);
            foreach (var (owner, filter) in plan.FastFilters.Concat(plan.ProcessFilters)) { Add(owner, filter); }
            if (options.HookTests || plan.HookTests) { Add("ScriptSafety", plan.ToolingFilters.GetValueOrDefault("ScriptSafety") ?? "RequiresExcel=false"); }
            if (options.SkillTests || plan.SkillTests) { Add("SkillGeneration", plan.ToolingFilters.GetValueOrDefault("SkillGeneration") ?? "Feature=SkillGeneration"); }
            if (options.PackagingTests || plan.PackagingTests) { Add("Packaging", plan.ToolingFilters.GetValueOrDefault("Packaging") ?? "RequiresExcel=false"); }
            if (options.Contracts) { foreach (var owner in new[] { "Core", "CLI", "McpServer" }) { Add(owner, "Feature=GeneratedContracts"); } }
        }
        else
        {
            foreach (var owner in new[] { "CLI", "ComInterop", "Core", "McpServer", "Service", "SkillGeneration", "Packaging", "ScriptSafety" }) { selections[owner] = "RequiresExcel=false"; }
        }
        return selections.Select(item => new FreeTestProject(item.Key, $"RequiresExcel=false&RunType!=OnDemand&({item.Value})")).ToArray();

        void Add(string owner, string filter) => selections[owner] = selections.TryGetValue(owner, out var previous)
            ? string.Join('|', previous.Split('|').Concat(filter.Split('|')).Distinct(StringComparer.Ordinal)) : filter;
    }

    public static async Task ExecuteAsync(string root, FreeTestOptions options, IProcessRunner runner, ValidationPlan? plan = null)
    {
        var selected = Select(root, options, plan);
        if (selected.Count == 0) { Console.Error.WriteLine("No local Excel-free tests selected."); return; }
        var results = options.ResultsDirectory ?? Path.Combine(root, "TestResults", $"excel-free-{Guid.NewGuid():N}");
        var execution = new TestExecution(root, runner);
        foreach (var item in selected) { await execution.RunAsync(item.Owner, item.Filter, results, excel: false, options.ListTests); }
    }
}
