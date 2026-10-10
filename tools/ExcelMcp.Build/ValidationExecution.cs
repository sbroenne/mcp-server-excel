namespace Sbroenne.ExcelMcp.Build;

public sealed class ValidationExecution(string root, IProcessRunner runner)
{
    private readonly TestExecution _tests = new(root, runner);

    public async Task BuildAsync(ValidationPlan plan, string? group = null)
    {
        var projects = BuildProjects(plan, group);
        foreach (var project in projects)
        {
            var path = Path.GetFullPath(project, root);
            if (!path.StartsWith(Path.TrimEndingDirectorySeparator(Path.GetFullPath(root)) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) ||
                (Path.GetExtension(path) != ".csproj" && Path.GetFileName(path) != "Sbroenne.ExcelMcp.sln"))
            {
                throw new ArgumentException($"Build input must be a repository project: {project}.");
            }
            if (!File.Exists(path)) { throw new InvalidOperationException($"Build input does not exist: {project}."); }
            Console.Error.WriteLine($"Building {project}.");
            var restore = await runner.CheckedAsync("dotnet", ["restore", path], TimeSpan.FromMinutes(20));
            Console.Error.WriteLine(restore.Output);
            if (!string.IsNullOrWhiteSpace(restore.Error)) { Console.Error.WriteLine(restore.Error); }
            var result = await runner.CheckedAsync("dotnet",
                ["build", path, "-c", "Release", "--no-restore", "--disable-build-servers", "-p:NuGetAudit=false", "--verbosity", "minimal"],
                TimeSpan.FromMinutes(20));
            Console.Error.WriteLine(result.Output);
            if (!string.IsNullOrWhiteSpace(result.Error)) { Console.Error.WriteLine(result.Error); }
        }
    }

    public string[] BuildProjects(ValidationPlan plan, string? group = null)
    {
        var owners = group is null or "Excel" ? [] : Owners(plan, group);
        var groups = plan.CiTestGroups;
        var fullBuildGroup = groups.Contains(plan.SourceChecksGroup, StringComparer.Ordinal)
            ? plan.SourceChecksGroup : groups.FirstOrDefault();
        var projects = plan.FullSolutionBuild && (group is null or "Excel" || group == fullBuildGroup)
            ? new[] { "Sbroenne.ExcelMcp.sln" }
            : group is null ? plan.BuildProjects.ToArray()
            : group == "Excel" ? plan.ExcelSelections.Select(selection => _tests.ProjectPath(selection.Project)).Distinct(StringComparer.Ordinal).ToArray()
            : owners.Select(owner => Path.GetRelativePath(root, _tests.ProjectPath(owner))).ToArray();
        if (group is not null && projects.Length == 0) { throw new InvalidOperationException($"{group} has no build projects."); }
        return projects;
    }

    public async Task CheckSourcesAsync(ValidationPlan plan)
    {
        if (plan.SourceChecks)
        {
            foreach (var script in new[] { "check-com-leaks", "check-success-flag", "check-dynamic-casts", "check-workbook-package-access" })
            {
                new SourceGuards(root).Check(script["check-".Length..]);
            }
        }
        if (plan.DocumentationCounts)
        {
            if (!plan.FullSolutionBuild) { throw new InvalidOperationException("Documentation counts require a complete Release build."); }
            foreach (var (script, option) in new[] { ("Build-AgentSkills", "-GenerateOnly"), ("check-doc-counts", "-SkipBuild") })
            {
                var result = await runner.CheckedAsync("pwsh", ["-NoProfile", "-File", Path.Combine(root, "scripts", $"{script}.ps1"), option],
                    TimeSpan.FromMinutes(10));
                Console.Error.WriteLine(result.Output);
            }
        }
    }

    public async Task TestAsync(ValidationPlan plan, string? group, string? resultsDirectory, bool listOnly = false, int? deadlineSeconds = null)
    {
        var results = resultsDirectory ?? Path.Combine(root, "TestResults", $"selected-{Guid.NewGuid():N}");
        foreach (var selected in group is null ? plan.CiTestGroups : group == "Excel" ? [] : [group])
        {
            var filters = Filters(plan, selected);
            if (filters.Count == 0) { throw new InvalidOperationException($"{selected} has no selected tests."); }
            foreach (var (owner, filter) in filters)
            {
                var kind = selected switch { "Fast" => "&AdapterTestKind!=System", "Process" => "&AdapterTestKind=System", _ => "" };
                await _tests.RunAsync(owner, $"RequiresExcel=false&RunType!=OnDemand{kind}&({filter})",
                    Path.Combine(results, selected), excel: false, listOnly, deadlineSeconds: deadlineSeconds);
            }
        }
        if (group is not null && group != "Excel") { return; }
        if (group == "Excel" && !plan.Excel) { throw new InvalidOperationException("The requested Excel group was not selected."); }
        if (!plan.Excel)
        {
            Console.Error.WriteLine("No Excel tests selected.");
            return;
        }
        if (!OperatingSystem.IsWindows() && !listOnly)
        {
            throw new InvalidOperationException("Selected Excel tests need Windows with desktop Excel. They were not run.");
        }
        var pipe = $"excelmcp-selected-{Environment.ProcessId}-{Guid.NewGuid():N}";
        try
        {
            foreach (var selection in plan.ExcelSelections.GroupBy(selection => selection.Project))
            {
                var filters = string.Join('|', selection.SelectMany(item => item.Filter.Split('|')).Distinct(StringComparer.Ordinal));
                var required = plan.FullE2E ? "&Acceptance!=Required" : "";
                await _tests.RunAsync(selection.Key, $"RequiresExcel=true&RunType!=OnDemand{required}&({filters})",
                    Path.Combine(results, "Excel"), excel: true, listOnly, pipe, deadlineSeconds);
            }
            if (plan.InfrastructureDiagnostics)
            {
                await _tests.RunAsync("ComInterop",
                    "RequiresExcel=true&RunType=OnDemand&FullyQualifiedName!~BeginBatch_RealIrmWorkbook&Locale!=ja-JP",
                    Path.Combine(results, "Infrastructure"), excel: true, listOnly, pipe, deadlineSeconds);
            }
            if (plan.FullE2E && !listOnly)
            {
                var result = await runner.CheckedAsync("pwsh",
                    ["-NoProfile", "-File", Path.Combine(root, "scripts", "Test-E2E.ps1"), "-SkipBuild", "-PipeName", pipe,
                     "-ResultsDirectory", Path.Combine(results, "Acceptance")], TimeSpan.FromHours(3));
                Console.Error.WriteLine(result.Output);
            }
        }
        catch (Exception primary)
        {
            try { await CleanupAsync(); }
            catch (Exception cleanup)
            {
                throw new AggregateException("Selected tests and owned cleanup failed.", primary, cleanup);
            }
            throw;
        }
        await CleanupAsync();

        async Task CleanupAsync()
        {
            if (listOnly || !plan.ExcelSelections.Any(selection => selection.Project == "CLI") && !plan.FullE2E) { return; }
            var result = await runner.CheckedAsync("pwsh",
                ["-NoProfile", "-File", Path.Combine(root, "scripts", "Stop-ExcelMcpProcesses.ps1"), "-PipeName", pipe],
                TimeSpan.FromMinutes(5));
            Console.Error.WriteLine(result.Output);
        }
    }

    internal static SortedDictionary<string, string> Filters(ValidationPlan plan, string group)
    {
        var owners = Owners(plan, group);
        var filters = group switch
        {
            "Fast" => plan.FastFilters,
            "Process" => plan.ProcessFilters,
            "Tooling" => plan.ToolingFilters,
            _ => throw new ArgumentException($"Unexpected test group: {group}.")
        };
        if (!owners.Order(StringComparer.Ordinal).SequenceEqual(filters.Keys, StringComparer.Ordinal)) { throw new InvalidOperationException($"Missing or unexpected {group} project filters."); }
        if (filters.Values.Any(string.IsNullOrWhiteSpace)) { throw new InvalidOperationException($"Missing {group} test filter."); }
        return filters;
    }

    private static string[] Owners(ValidationPlan plan, string group)
    {
        if (!plan.CiTestGroups.Contains(group, StringComparer.Ordinal)) { throw new InvalidOperationException($"The requested {group} group was not selected."); }
        var owners = group switch
        {
            "Fast" => plan.FastProjects,
            "Process" => plan.ProcessProjects,
            "Tooling" => plan.ToolingProjects,
            _ => throw new ArgumentException($"Unexpected test group: {group}.")
        };
        foreach (var owner in owners)
        {
            var allowed = group switch
            {
                "Fast" => owner is "CLI" or "ComInterop" or "Core" or "McpServer" or "Service" or "Diagnostics",
                "Process" => owner == "CLI",
                _ => owner is "Packaging" or "ScriptSafety" or "SkillGeneration"
            };
            if (!allowed) { throw new ArgumentException($"Unexpected {group} project: {owner}."); }
        }
        return [.. owners];
    }
}
