using System.Text.Json;

namespace Sbroenne.ExcelMcp.Build;

public sealed class ExcelGroupOptions
{
    public string[] Groups { get; set; } = ["Editing", "Reporting", "Data", "Lifecycle", "Infrastructure", "Acceptance", "VBA", "Desktop"];
    public bool IncludeInfrastructureDiagnostics { get; set; }
    public bool ListTests { get; set; }
    public string? ResultsDirectory { get; set; }
    public int DeadlineSeconds { get; set; } = 7200;
}

public sealed record BuiltTestCase(string Class, string Method, string Group, bool OnDemand, bool Prerequisite, bool Required);
public sealed record BuiltTestGroup(string Project, string ProjectName, string Class, string Group, string Filter, BuiltTestCase[] Cases);

public sealed class ExcelGroupExecution(string root, IProcessRunner runner)
{
    private readonly TestExecution _tests = new(root, runner);
    private static readonly HashSet<string> Groups = new(StringComparer.Ordinal) { "Editing", "Reporting", "Data", "Lifecycle", "Infrastructure", "Acceptance", "VBA", "Desktop" };

    public async Task<BuiltTestGroup[]> InventoryAsync(string results)
    {
        results = Path.GetFullPath(results, root);
        var inventory = new List<BuiltTestGroup>();
        foreach (var owner in new[] { "Core", "Service", "CLI", "McpServer", "ComInterop" })
        {
            var path = Path.Combine(results, $"{owner}-inventory.json");
            if (File.Exists(path)) { throw new InvalidOperationException($"Use a new inventory path: {path}"); }
            await _tests.RunStageAsync(new TestStageOptions
            {
                Project = _tests.ProjectPath(owner),
                Filter = "FullyQualifiedName~ExcelValidationGroups_PreserveClassFixturesAndExportBuiltInventory",
                Name = $"{owner}-inventory",
                ResultsDirectory = results,
                DeadlineSeconds = 120,
                Environment = new Dictionary<string, string> { ["EXCELMCP_TEST_SELECTION_OUTPUT"] = path }
            });
            if (!File.Exists(path)) { throw new InvalidOperationException($"Missing built test inventory: {path}"); }
            var cases = JsonSerializer.Deserialize<BuiltTestCase[]>(await File.ReadAllTextAsync(path))
                ?? throw new InvalidOperationException($"Empty built test inventory: {path}");
            foreach (var type in cases.GroupBy(item => item.Class))
            {
                var groups = type.Select(item => item.Group).Distinct(StringComparer.Ordinal).ToArray();
                foreach (var group in groups)
                {
                    var selected = type.Where(item => item.Group == group).ToArray();
                    var filter = groups.Length == 1 ? $"FullyQualifiedName~{type.Key}." : string.Join('|', selected.Select(item => $"FullyQualifiedName={item.Method}"));
                    inventory.Add(new BuiltTestGroup(_tests.ProjectPath(owner), owner, type.Key, group, $"({filter})", selected));
                }
            }
        }
        return [.. inventory];
    }

    public async Task ExecuteAsync(ExcelGroupOptions options)
    {
        if (options.Groups.Length == 0 || options.Groups.Any(group => !Groups.Contains(group))) { throw new ArgumentException("Select one or more known Excel groups."); }
        if (!OperatingSystem.IsWindows()) { throw new InvalidOperationException("Real-Excel validation requires Windows with desktop Excel."); }
        if (!options.ListTests && Type.GetTypeFromProgID("Excel.Application") is null) { throw new InvalidOperationException("Desktop Excel is not registered. Selected Excel tests were not run."); }
        var results = options.ResultsDirectory ?? Path.Combine(root, "TestResults", $"excel-{Guid.NewGuid():N}");
        var inventory = await InventoryAsync(Path.Combine(results, "inventory"));
        var pipe = $"excelmcp-feature-{Environment.ProcessId}-{Guid.NewGuid():N}";
        Exception? primary = null;
        try
        {
            foreach (var group in options.Groups.Distinct(StringComparer.Ordinal))
            {
                var selected = inventory.Where(item => item.Group == group && item.Cases.Any(test => !test.OnDemand)).ToArray();
                if (selected.Length == 0) { throw new InvalidOperationException($"{group} has no matching Excel classes."); }
                if (group == "Acceptance" && !options.ListTests)
                {
                    var result = await runner.CheckedAsync("pwsh", ["-NoProfile", "-File", Path.Combine(root, "scripts", "Test-E2E.ps1"),
                        "-SkipBuild", "-PipeName", pipe, "-ResultsDirectory", Path.Combine(results, "acceptance")], TimeSpan.FromHours(3));
                    Console.Error.WriteLine(result.Output);
                    selected = selected.Where(item => item.Cases.Any(test => !test.Required && !test.OnDemand))
                        .Select(item => item with { Filter = "(" + string.Join('|', item.Cases.Where(test => !test.Required && !test.OnDemand).Select(test => $"FullyQualifiedName={test.Method}")) + ")" }).ToArray();
                }
                foreach (var project in selected.GroupBy(item => item.ProjectName))
                {
                    await _tests.RunStageAsync(new TestStageOptions
                    {
                        Project = _tests.ProjectPath(project.Key),
                        Filter = $"RequiresExcel=true&RunType!=OnDemand&({string.Join('|', project.Select(item => item.Filter))})",
                        ResultsDirectory = results,
                        Name = $"{group}-{project.Key}",
                        DeadlineSeconds = options.DeadlineSeconds,
                        HangTimeout = "10m",
                        ListTests = options.ListTests,
                        Environment = new Dictionary<string, string> { ["EXCELMCP_CLI_PIPE"] = pipe }
                    });
                }
            }
            if (options.IncludeInfrastructureDiagnostics)
            {
                await _tests.RunStageAsync(new TestStageOptions
                {
                    Project = _tests.ProjectPath("ComInterop"),
                    Filter = "RequiresExcel=true&RunType=OnDemand&FullyQualifiedName!~BeginBatch_RealIrmWorkbook&Locale!=ja-JP",
                    ResultsDirectory = results,
                    Name = "Infrastructure-OnDemand",
                    DeadlineSeconds = 5400,
                    HangTimeout = "10m",
                    ListTests = options.ListTests,
                    Environment = new Dictionary<string, string> { ["EXCELMCP_CLI_PIPE"] = pipe }
                });
            }
        }
        catch (Exception exception) { primary = exception; }
        try
        {
            if (!options.ListTests)
            {
                var result = await runner.CheckedAsync("pwsh", ["-NoProfile", "-File", Path.Combine(root, "scripts", "Stop-ExcelMcpProcesses.ps1"), "-PipeName", pipe], TimeSpan.FromMinutes(5));
                Console.Error.WriteLine(result.Output);
            }
        }
        catch (Exception cleanup) when (cleanup is InvalidOperationException or IOException or TimeoutException)
        {
            if (primary is not null) { throw new AggregateException("Excel validation and owned cleanup failed.", primary, cleanup); }
            throw;
        }
        if (primary is not null) { System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(primary).Throw(); }
    }
}
