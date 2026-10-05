using System.Globalization;
using System.Reflection;
using System.Xml.Linq;

namespace Sbroenne.ExcelMcp.Build;

public static class TestReport
{
    public static int Verify(string path)
    {
        if (!File.Exists(path)) { throw new InvalidOperationException($"Missing test report: {path}."); }
        var document = XDocument.Load(path);
        var summary = document.Descendants().SingleOrDefault(element => element.Name.LocalName == "ResultSummary");
        var counters = summary?.Elements().SingleOrDefault(element => element.Name.LocalName == "Counters");
        var total = Counter(counters, "total");
        var passed = Counter(counters, "passed");
        var executed = Counter(counters, "executed");
        var results = document.Descendants().Where(element => element.Name.LocalName == "UnitTestResult").ToArray();
        if (total <= 0 || passed != total || executed != total ||
            summary?.Attribute("outcome")?.Value != "Completed" ||
            results.Length != total || results.Any(result => result.Attribute("outcome")?.Value != "Passed") ||
            summary.Descendants().Any(element => element.Name.LocalName == "RunInfo" && element.Attribute("outcome")?.Value == "Error"))
        {
            throw new InvalidOperationException($"Test selection was empty, skipped, failed, or had a cleanup failure. See {path}.");
        }
        return total;
    }

    private static int Counter(XElement? counters, string name) =>
        int.TryParse(counters?.Attribute(name)?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var value) ? value : -1;
}

public sealed class TestStageOptions
{
    public string Project { get; set; } = "";
    public string Filter { get; set; } = "";
    public string ResultsDirectory { get; set; } = "";
    public string Name { get; set; } = "";
    public int DeadlineSeconds { get; set; } = 1800;
    public string HangTimeout { get; set; } = "5m";
    public Dictionary<string, string> Environment { get; set; } = new(StringComparer.Ordinal);
    public bool ListTests { get; set; }
}

public sealed class TestExecution(string root, IProcessRunner runner)
{
    public async Task RunAsync(string owner, string filter, string results, bool excel, bool listOnly = false, string? pipe = null, int? deadlineSeconds = null)
    {
        var options = new TestStageOptions
        {
            Project = ProjectPath(owner),
            Filter = filter,
            ResultsDirectory = results,
            Name = owner,
            DeadlineSeconds = deadlineSeconds ?? (excel ? 7200 : 1800),
            HangTimeout = excel ? "10m" : "5m",
            ListTests = listOnly
        };
        if (pipe is not null) { options.Environment["EXCELMCP_CLI_PIPE"] = pipe; }
        await RunStageAsync(options);
    }

    public async Task RunStageAsync(TestStageOptions options)
    {
        var name = options.Name;
        if (string.IsNullOrWhiteSpace(options.Filter)) { throw new ArgumentException($"Missing filter for {name}."); }
        if (name.Length == 0 || name.Any(character => !char.IsAsciiLetterOrDigit(character) && character is not ('-' or '_')))
        {
            throw new ArgumentException("A test-stage name must contain only letters, numbers, underscores or hyphens.");
        }
        if (options.DeadlineSeconds is < 1 or > 28800) { throw new ArgumentException("A hard deadline must be 1-28800 seconds."); }
        var project = Path.GetFullPath(options.Project, root);
        if (!File.Exists(project) || Path.GetExtension(project) is not (".csproj" or ".proj")) { throw new ArgumentException($"Missing or invalid test project: {project}."); }
        var results = Path.GetFullPath(options.ResultsDirectory, root);
        Directory.CreateDirectory(results);
        var report = Path.Combine(results, $"{name}.trx");
        if (File.Exists(report)) { throw new InvalidOperationException($"Use a new report path: {report}."); }
        var arguments = new List<string>
        {
            "test", project, "-c", "Release", "--no-build", "--no-restore", "--disable-build-servers",
            "--filter", options.Filter, "--blame-hang-timeout", options.HangTimeout,
            "--results-directory", results, "--logger", $"trx;LogFileName={name}.trx"
        };
        if (options.ListTests) { arguments.Add("--list-tests"); }
        Console.Error.WriteLine($"{name} : {options.Filter}");
        var environment = new Dictionary<string, string>
        {
            ["EXCELMCP_TEST_OWNERSHIP_DIRECTORY"] = Path.Combine(results, $"{name}-ownership"),
            ["EXCELMCP_BUILD_DLL"] = Assembly.GetExecutingAssembly().Location,
            ["EXCELMCP_BUILD_ROOT"] = root
        };
        foreach (var (key, value) in options.Environment) { environment[key] = value; }
        ProcessResult result;
        try
        {
            result = await runner.RunAsync("dotnet", arguments,
                TimeSpan.FromSeconds(options.DeadlineSeconds), environment);
        }
        catch (TimeoutException exception)
        {
            await File.WriteAllTextAsync(Path.Combine(results, $"{name}.deadline.log"), exception.Message);
            throw new TimeoutException($"{name} exceeded the {options.DeadlineSeconds}-second hard deadline. Evidence: {results}.\n{exception.Message}", exception);
        }
        await File.WriteAllTextAsync(Path.Combine(results, $"{name}.stdout.log"), result.Output);
        await File.WriteAllTextAsync(Path.Combine(results, $"{name}.stderr.log"), result.Error);
        Console.Error.WriteLine(result.Output);
        if (!string.IsNullOrWhiteSpace(result.Error)) { Console.Error.WriteLine(result.Error); }
        if (result.ExitCode != 0) { throw new InvalidOperationException($"{name} failed with exit code {result.ExitCode}. Evidence: {results}.\n{result.Output}\n{result.Error}"); }
        if (!options.ListTests) { Console.Error.WriteLine($"{name}: {TestReport.Verify(report)} cases passed."); }
    }

    public string ProjectPath(string owner)
    {
        if (owner is not ("Core" or "ComInterop" or "CLI" or "McpServer" or "Service" or "Diagnostics" or "Packaging" or "ScriptSafety" or "SkillGeneration"))
        {
            throw new ArgumentException($"Unexpected test project: {owner}.");
        }
        var project = Path.Combine(root, "tests", $"ExcelMcp.{owner}.Tests", $"ExcelMcp.{owner}.Tests.csproj");
        if (!File.Exists(project)) { throw new InvalidOperationException($"Missing test project: {owner}."); }
        return project;
    }
}
