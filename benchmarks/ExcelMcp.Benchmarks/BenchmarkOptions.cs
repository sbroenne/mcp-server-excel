namespace Sbroenne.ExcelMcp.Benchmarks;

internal enum BenchmarkCommand
{
    Run,
    Compare,
    Catalog
}

internal sealed record BenchmarkOptions(
    BenchmarkCommand Command,
    string Profile,
    BenchmarkConfiguration Configuration,
    string OutputDirectory,
    string RepoRoot,
    IReadOnlySet<string> SelectedPlans,
    string? BaselinePath,
    string? CandidatePath,
    TimeSpan MaximumRunTime)
{
    public static BenchmarkOptions Parse(string[] args)
    {
        ArgumentNullException.ThrowIfNull(args);
        var command = args.Length > 0 && !args[0].StartsWith("--", StringComparison.Ordinal)
            ? args[0].ToLowerInvariant() switch
            {
                "run" => BenchmarkCommand.Run,
                "compare" => BenchmarkCommand.Compare,
                "catalog" => BenchmarkCommand.Catalog,
                _ => throw new ArgumentException($"Unknown command '{args[0]}'. Use run, compare, or catalog.")
            }
            : BenchmarkCommand.Run;

        var startIndex = args.Length > 0 && !args[0].StartsWith("--", StringComparison.Ordinal) ? 1 : 0;
        var values = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        for (var index = startIndex; index < args.Length; index++)
        {
            var token = args[index];
            if (!token.StartsWith("--", StringComparison.Ordinal))
            {
                throw new ArgumentException($"Unexpected argument '{token}'. Options must start with --.");
            }

            var key = token[2..];
            if (key is "show" or "hidden")
            {
                values[key] = "true";
                continue;
            }

            if (index + 1 >= args.Length || args[index + 1].StartsWith("--", StringComparison.Ordinal))
            {
                throw new ArgumentException($"Option '--{key}' requires a value.");
            }

            values[key] = args[++index];
        }

        var profile = Get(values, "profile")?.ToLowerInvariant() ?? "standard";
        var defaults = profile switch
        {
            "quick" => new BenchmarkConfiguration(1, 3, 3, false),
            "standard" => new BenchmarkConfiguration(3, 10, 20, false),
            "reliable" => new BenchmarkConfiguration(5, 30, 100, false),
            _ => throw new ArgumentException("Profile must be quick, standard, or reliable.")
        };

        var showExcel = values.ContainsKey("show") && !values.ContainsKey("hidden");
        var configuration = defaults with
        {
            Warmups = ParsePositiveOrZero(values, "warmups", defaults.Warmups),
            Iterations = ParsePositive(values, "iterations", defaults.Iterations),
            ReliabilityIterations = ParsePositive(values, "reliability-iterations", defaults.ReliabilityIterations),
            ShowExcel = showExcel
        };

        var repoRoot = Path.GetFullPath(Get(values, "repo") ?? Directory.GetCurrentDirectory());
        var runId = DateTimeOffset.UtcNow.ToString("yyyyMMdd-HHmmss", System.Globalization.CultureInfo.InvariantCulture);
        var output = Path.GetFullPath(Get(values, "output") ?? Path.Combine(repoRoot, "artifacts", "benchmarks", runId));
        var selectedPlans = (Get(values, "plans") ?? string.Join(',', BenchmarkPlanCatalog.All.Select(plan => plan.Id)))
            .Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
            .ToHashSet(StringComparer.Ordinal);

        var unknownPlans = selectedPlans.Except(BenchmarkPlanCatalog.All.Select(plan => plan.Id), StringComparer.Ordinal).ToArray();
        if (unknownPlans.Length > 0)
        {
            throw new ArgumentException($"Unknown plan IDs: {string.Join(", ", unknownPlans)}");
        }

        var maximumMinutes = ParsePositive(values, "maximum-minutes", profile == "reliable" ? 180 : 90);
        var baselinePath = Get(values, "baseline");
        var candidatePath = Get(values, "candidate");
        if (command == BenchmarkCommand.Compare &&
            (string.IsNullOrWhiteSpace(baselinePath) || string.IsNullOrWhiteSpace(candidatePath)))
        {
            throw new ArgumentException("compare requires --baseline <baseline.json> and --candidate <candidate.json>.");
        }

        return new BenchmarkOptions(
            command,
            profile,
            configuration,
            output,
            repoRoot,
            selectedPlans,
            baselinePath is null ? null : Path.GetFullPath(baselinePath),
            candidatePath is null ? null : Path.GetFullPath(candidatePath),
            TimeSpan.FromMinutes(maximumMinutes));
    }

    private static string? Get(Dictionary<string, string?> values, string key) =>
        values.TryGetValue(key, out var value) ? value : null;

    private static int ParsePositive(Dictionary<string, string?> values, string key, int fallback)
    {
        if (!values.TryGetValue(key, out var raw))
        {
            return fallback;
        }

        if (!int.TryParse(raw, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var value) || value < 1)
        {
            throw new ArgumentException($"--{key} must be a positive integer.");
        }

        return value;
    }

    private static int ParsePositiveOrZero(Dictionary<string, string?> values, string key, int fallback)
    {
        if (!values.TryGetValue(key, out var raw))
        {
            return fallback;
        }

        if (!int.TryParse(raw, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var value) || value < 0)
        {
            throw new ArgumentException($"--{key} must be zero or a positive integer.");
        }

        return value;
    }
}
