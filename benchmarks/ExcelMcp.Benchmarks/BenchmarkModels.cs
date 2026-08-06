namespace Sbroenne.ExcelMcp.Benchmarks;

internal sealed record BenchmarkConfiguration(
    int Warmups,
    int Iterations,
    int ReliabilityIterations,
    bool ShowExcel);

internal sealed record BenchmarkEnvironment(
    string MachineNameHash,
    string DotNetVersion,
    string OperatingSystem,
    string ExcelVersion,
    string ProcessArchitecture,
    int LogicalProcessors,
    long TotalMemoryBytes,
    string GitCommit,
    string GitBranch,
    bool GitDirty);

internal sealed record BenchmarkObservation(
    int Iteration,
    string Case,
    bool Success,
    string? Error,
    IReadOnlyDictionary<string, double> Metrics,
    IReadOnlyDictionary<string, string>? Dimensions = null,
    string? Outcome = null);

internal sealed record BenchmarkInvariant(string Name, bool Passed, string Evidence);

internal sealed record ScenarioResult(
    string PlanId,
    string Scenario,
    string Title,
    string Workload,
    string Status,
    IReadOnlyList<BenchmarkObservation> Observations,
    IReadOnlyList<BenchmarkInvariant> Invariants,
    IReadOnlyDictionary<string, DistributionSummary> Summaries,
    IReadOnlyDictionary<string, IReadOnlyDictionary<string, DistributionSummary>> CaseSummaries,
    ReliabilitySummary Reliability,
    IReadOnlyList<string> Notes)
{
    public static ScenarioResult Create(
        string planId,
        string scenario,
        string title,
        IReadOnlyList<BenchmarkObservation> observations,
        IReadOnlyList<BenchmarkInvariant> invariants,
        string workload,
        IReadOnlyList<string>? notes = null,
        string? status = null)
    {
        ArgumentNullException.ThrowIfNull(observations);
        ArgumentNullException.ThrowIfNull(invariants);

        var summaries = SummarizeMetrics(observations);
        var caseSummaries = observations
            .GroupBy(observation => observation.Case, StringComparer.Ordinal)
            .OrderBy(group => group.Key, StringComparer.Ordinal)
            .ToDictionary(
                group => group.Key,
                group => (IReadOnlyDictionary<string, DistributionSummary>)SummarizeMetrics(group),
                StringComparer.Ordinal);

        var successes = observations.Count(observation => observation.Success);
        var failures = observations.Count - successes;
        var reliability = Statistics.SummarizeReliability(successes, failures);
        var effectiveStatus = status ?? (invariants.All(invariant => invariant.Passed) && failures == 0 ? "passed" : "failed");

        return new ScenarioResult(
            planId,
            scenario,
            title,
            workload,
            effectiveStatus,
            observations,
            invariants,
            summaries,
            caseSummaries,
            reliability,
            notes ?? []);
    }

    private static Dictionary<string, DistributionSummary> SummarizeMetrics(
        IEnumerable<BenchmarkObservation> observations) =>
        observations
            .SelectMany(observation => observation.Metrics)
            .Where(metric => double.IsFinite(metric.Value))
            .GroupBy(metric => metric.Key, StringComparer.Ordinal)
            .OrderBy(group => group.Key, StringComparer.Ordinal)
            .ToDictionary(
                group => group.Key,
                group => Statistics.Summarize(group.Select(metric => metric.Value).ToArray()),
                StringComparer.Ordinal);
}

internal sealed record BenchmarkRunReport(
    string RunId,
    string Profile,
    DateTimeOffset StartedAtUtc,
    DateTimeOffset FinishedAtUtc,
    BenchmarkConfiguration Configuration,
    BenchmarkEnvironment Environment,
    IReadOnlyList<ScenarioResult> Scenarios,
    string SchemaVersion = "1.0",
    string RunState = "completed");

internal sealed record ReportPaths(string JsonPath, string CsvPath, string MarkdownPath);
