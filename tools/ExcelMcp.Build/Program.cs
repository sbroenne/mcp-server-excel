using System.Text.Json;
using System.Text.Json.Serialization;
using System.Xml;

namespace Sbroenne.ExcelMcp.Build;

internal static class Program
{
    internal static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PreferredObjectCreationHandling = JsonObjectCreationHandling.Populate
    };

    private static async Task<int> Main(string[] args)
    {
        try
        {
            if (args.Length == 0 || args[0] is "help" or "--help")
            {
                Console.WriteLine("ExcelMcp development tooling: plan, build, test, validate, check-source, package, test-free, test-excel, inventory, stage, verify-report, complete.");
                Console.WriteLine("Use plan --base-ref <ref> or --staged; --full explicitly selects complete validation.");
                return 0;
            }
            var options = new CommandOptions(args[1..]);
            var root = options.Get("--root") ?? FindRoot();
            if (!File.Exists(Path.Combine(root, "Sbroenne.ExcelMcp.sln")))
            {
                throw new ArgumentException("The build root must contain Sbroenne.ExcelMcp.sln.");
            }
            var runner = new ProcessRunner(root);
            switch (args[0])
            {
                case "verify-report":
                    TestReport.Verify(options.Get("--report") ?? throw new ArgumentException("Select --report."));
                    return 0;
                case "stage":
                    await new TestExecution(root, runner).RunStageAsync(await ReadOptionsAsync<TestStageOptions>(options, "--stage-options"));
                    return 0;
                case "test-free":
                    var free = await ReadOptionsAsync<FreeTestOptions>(options, "--test-options");
                    var saved = free.PlanFile is null ? null : JsonSerializer.Deserialize<ValidationPlan>(await File.ReadAllTextAsync(free.PlanFile), JsonOptions)
                        ?? throw new ArgumentException("The validation plan is empty.");
                    if (saved is not null && saved.SchemaVersion != 1) { throw new ArgumentException("Unsupported validation-plan version. Generate a fresh plan."); }
                    await FreeTestSelection.ExecuteAsync(root, free, runner, saved);
                    return 0;
                case "test-excel":
                    await new ExcelGroupExecution(root, runner).ExecuteAsync(await ReadOptionsAsync<ExcelGroupOptions>(options, "--test-options"));
                    return 0;
                case "inventory":
                    Console.WriteLine(JsonSerializer.Serialize(await new ExcelGroupExecution(root, runner)
                        .InventoryAsync(options.Get("--results-directory") ?? throw new ArgumentException("Select --results-directory.")), JsonOptions));
                    return 0;
                case "complete":
                    CiCompletion.Verify(await ReadOptionsAsync<CiCompletionOptions>(options, "--completion-options"));
                    return 0;
                case "package":
                    {
                        var package = options.Get("--package-options") is { } file
                            ? JsonSerializer.Deserialize<PackageOptions>(await File.ReadAllTextAsync(file), JsonOptions) ?? throw new ArgumentException("Package options are empty.")
                            : new PackageOptions
                            {
                                Components = (options.Get("--components") ?? "Cli,Mcp,Extension,Mcpb,Skills,Plugins").Split(',', StringSplitOptions.RemoveEmptyEntries),
                                Version = options.Get("--version"),
                                OutputDirectory = options.Get("--output"),
                                SkillsDirectory = options.Get("--skills-directory"),
                                McpRuntimeExecutable = options.Get("--mcp-runtime"),
                                CliRuntimeExecutable = options.Get("--cli-runtime"),
                                SkipExtensionTests = options.Has("--skip-extension-tests"),
                                BaseRef = options.Get("--base-ref"),
                                HeadRef = options.Get("--head-ref") ?? "HEAD"
                            };
                        await new PackageExecution(Path.GetFullPath(options.Get("--source-root") ?? root)).ExecuteAsync(package);
                        return 0;
                    }
                case "check-source":
                    {
                        var inputs = options.Get("--inputs-json") is { } inputJson
                            ? JsonSerializer.Deserialize<string[]>(inputJson) ?? throw new ArgumentException("Inputs must be a JSON array.")
                            : null;
                        new SourceGuards(Path.GetFullPath(options.Get("--scan-root") ?? root))
                            .Check(options.Get("--rule") ?? throw new ArgumentException("Select --rule."), inputs);
                        return 0;
                    }
                case "plan":
                    {
                        var paths = options.Get("--paths-json") is { } json
                            ? JsonSerializer.Deserialize<string[]>(json) ?? throw new ArgumentException("Paths must be a JSON array.")
                            : options.Has("--full") ? [] : await ReadChangedPathsAsync(runner, options);
                        var plan = new ValidationPolicy(root).Select(paths, options.Has("--full"));
                        var content = JsonSerializer.Serialize(plan, JsonOptions);
                        if (options.Get("--output") is { } destination)
                        {
                            destination = Path.GetFullPath(destination, root);
                            Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
                            await File.WriteAllTextAsync(destination, content);
                        }
                        Console.WriteLine(content);
                        return 0;
                    }
                case "build":
                case "test":
                case "validate":
                    {
                        var plan = await ReadPlanAsync(root, options);
                        var execution = new ValidationExecution(root, runner);
                        if (args[0] == "build" && options.Has("--list-projects"))
                        {
                            foreach (var project in execution.BuildProjects(plan, options.Get("--group"))) { Console.WriteLine(project); }
                            return 0;
                        }
                        if (args[0] != "test" || !options.Has("--no-build"))
                        {
                            await execution.BuildAsync(plan, options.Get("--group"));
                        }

                        if (args[0] == "validate")
                        {
                            await execution.CheckSourcesAsync(plan);
                        }
                        if (args[0] != "build")
                        {
                            await execution.TestAsync(plan, options.Get("--group"), options.Get("--results-directory"), options.Has("--list-tests"), options.DeadlineSeconds());
                        }
                        return 0;
                    }
                default:
                    throw new ArgumentException($"Unknown build target '{args[0]}'.");
            }
        }
        catch (Exception exception) when (exception is ArgumentException or InvalidOperationException or IOException or UnauthorizedAccessException or TimeoutException or JsonException or XmlException or AggregateException)
        {
            Console.Error.WriteLine(exception is AggregateException aggregate
                ? aggregate.Message + Environment.NewLine + string.Join(Environment.NewLine, aggregate.Flatten().InnerExceptions.Select(error => error.Message))
                : exception.Message);
            return 1;
        }
    }

    private static async Task<T> ReadOptionsAsync<T>(CommandOptions options, string name) where T : class =>
        JsonSerializer.Deserialize<T>(await File.ReadAllTextAsync(options.Get(name) ?? throw new ArgumentException($"Select {name}.")), JsonOptions)
        ?? throw new ArgumentException($"Empty command options: {name}");

    private static async Task<ValidationPlan> ReadPlanAsync(string root, CommandOptions options)
    {
        if (options.Get("--plan") is { } file)
        {
            var plan = JsonSerializer.Deserialize<ValidationPlan>(await File.ReadAllTextAsync(Path.GetFullPath(file, root)), JsonOptions)
                ?? throw new ArgumentException("The validation plan is empty.");
            if (plan.SchemaVersion != 1) { throw new ArgumentException("Unsupported validation-plan version. Generate a fresh plan."); }
            if (plan.Excel && plan.ExcelSelections.Count == 0 && !plan.FullE2E)
            {
                throw new ArgumentException("Excel was selected but the plan contains no test filters.");
            }
            return plan;
        }
        if (options.Get("--area") is { } area)
        {
            if (area.Contains(Path.DirectorySeparatorChar) || area.Contains('/') || area.Contains("..", StringComparison.Ordinal))
            {
                throw new ArgumentException("An area must be a single command category.");
            }
            return new ValidationPolicy(root).Select([$"src/ExcelMcp.Core/Commands/{area}/Commands.cs"]);
        }
        if (options.Has("--full")) { return new ValidationPolicy(root).Select([], full: true); }
        if (options.Has("--staged") || options.Get("--base-ref") is not null)
        {
            return new ValidationPolicy(root).Select(await ReadChangedPathsAsync(new ProcessRunner(root), options));
        }
        throw new ArgumentException("Select --plan, --area, --staged, --base-ref, or --full.");
    }

    private static async Task<string[]> ReadChangedPathsAsync(ProcessRunner runner, CommandOptions options)
    {
        var arguments = new List<string> { "-c", "core.quotepath=false", "diff", "--name-only", "--no-renames" };
        if (options.Has("--staged"))
        {
            var merge = await runner.RunAsync("git", ["rev-parse", "--verify", "--quiet", "MERGE_HEAD"], TimeSpan.FromSeconds(30), preserveGitContext: true);
            arguments.Add("--cached");
            arguments.Add(merge.ExitCode == 0 ? merge.Output.Trim() : "HEAD");
        }
        else
        {
            var baseRef = options.Get("--base-ref") ?? throw new ArgumentException("Select --base-ref, --staged, --paths-json, or --full.");
            var head = options.Get("--head-ref") ?? "HEAD";
            arguments.Add($"{baseRef}...{head}");
        }
        var result = await runner.CheckedAsync("git", arguments, TimeSpan.FromSeconds(30), preserveGitContext: options.Has("--staged"));
        return result.Output.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
    }

    private static string FindRoot()
    {
        for (var directory = new DirectoryInfo(Environment.CurrentDirectory); directory is not null; directory = directory.Parent)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Sbroenne.ExcelMcp.sln"))) { return directory.FullName; }
        }
        throw new DirectoryNotFoundException("Cannot find the repository root.");
    }
}

internal sealed class CommandOptions
{
    private readonly Dictionary<string, string?> _values = new(StringComparer.Ordinal);
    private static readonly HashSet<string> Flags = new(StringComparer.Ordinal) { "--full", "--staged", "--no-build", "--list-tests", "--skip-extension-tests", "--list-projects" };
    private static readonly HashSet<string> Values = new(StringComparer.Ordinal) {
        "--root", "--paths-json", "--base-ref", "--head-ref", "--output", "--plan", "--group", "--area",
        "--results-directory", "--deadline-seconds", "--rule", "--scan-root", "--inputs-json", "--package-options", "--source-root",
        "--components", "--version", "--skills-directory", "--mcp-runtime", "--cli-runtime",
        "--test-options", "--stage-options", "--completion-options", "--report"
    };

    public CommandOptions(string[] arguments)
    {
        for (var index = 0; index < arguments.Length; index++)
        {
            var name = arguments[index];
            if (!_values.TryAdd(name, null)) { throw new ArgumentException($"Duplicate option {name}."); }
            if (Flags.Contains(name)) { continue; }
            if (!Values.Contains(name) || index + 1 >= arguments.Length || arguments[index + 1].StartsWith("--", StringComparison.Ordinal))
            {
                throw new ArgumentException($"Unknown option or missing value: {name}.");
            }
            _values[name] = arguments[++index];
        }
    }

    public bool Has(string name) => _values.ContainsKey(name);
    public string? Get(string name) => _values.GetValueOrDefault(name);

    public int? DeadlineSeconds()
    {
        if (Get("--deadline-seconds") is not { } value) { return null; }
        if (!int.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var seconds) ||
            seconds is < 1 or > 28800)
        {
            throw new ArgumentException("A hard deadline must be 1-28800 seconds.");
        }
        return seconds;
    }
}
