namespace Sbroenne.ExcelMcp.Benchmarks;

internal sealed record MetricImpact(
    string Case,
    string Metric,
    bool LowerIsBetter,
    double BaselineMedian,
    double CandidateMedian,
    double BaselineP95,
    double CandidateP95,
    double CandidateToBaselineRatio,
    double PercentImprovement,
    bool Improved);

internal sealed record ScenarioImpact(
    string PlanId,
    string Scenario,
    string Title,
    bool SafetyInvariantsPassed,
    IReadOnlyDictionary<string, MetricImpact> Metrics);

internal sealed record BenchmarkComparisonReport(
    string BaselineRunId,
    string CandidateRunId,
    IReadOnlyList<ScenarioImpact> Scenarios);

internal static class ComparisonEngine
{
    private static readonly HashSet<string> HigherIsBetterMetrics = new(StringComparer.Ordinal)
    {
        "operations_per_second",
        "cells_per_second",
        "rows_per_second",
        "stale_detection_rate",
        "cleanup_success_rate"
    };

    public static BenchmarkComparisonReport Compare(BenchmarkRunReport baseline, BenchmarkRunReport candidate)
    {
        ArgumentNullException.ThrowIfNull(baseline);
        ArgumentNullException.ThrowIfNull(candidate);

        if (!string.Equals(baseline.RunState, "completed", StringComparison.Ordinal) ||
            !string.Equals(candidate.RunState, "completed", StringComparison.Ordinal))
        {
            throw new InvalidDataException("Both reports must have runState 'completed' before comparison.");
        }

        if (!string.Equals(baseline.Profile, candidate.Profile, StringComparison.Ordinal) ||
            baseline.Configuration != candidate.Configuration)
        {
            throw new InvalidDataException(
                "A one-to-one comparison requires the same profile and benchmark configuration, including visibility and repetition counts.");
        }

        if (!SameExecutionEnvironment(baseline.Environment, candidate.Environment))
        {
            throw new InvalidDataException(
                "A one-to-one comparison requires the same machine, Excel version, OS, architecture, and logical processor count.");
        }

        var candidateByScenario = candidate.Scenarios.ToDictionary(item => item.Scenario, StringComparer.Ordinal);
        var baselineScenarioNames = baseline.Scenarios.Select(item => item.Scenario).ToHashSet(StringComparer.Ordinal);
        var missingScenarios = baselineScenarioNames.Except(candidateByScenario.Keys, StringComparer.Ordinal).ToArray();
        var unexpectedScenarios = candidateByScenario.Keys.Except(baselineScenarioNames, StringComparer.Ordinal).ToArray();
        if (missingScenarios.Length > 0 || unexpectedScenarios.Length > 0)
        {
            throw new InvalidDataException(
                $"A one-to-one comparison requires identical scenario coverage. " +
                $"Missing candidate scenarios: {FormatNames(missingScenarios)}. " +
                $"Unexpected candidate scenarios: {FormatNames(unexpectedScenarios)}.");
        }

        var comparisons = new List<ScenarioImpact>();

        foreach (var baselineScenario in baseline.Scenarios.OrderBy(item => item.PlanId, StringComparer.Ordinal))
        {
            var candidateScenario = candidateByScenario[baselineScenario.Scenario];

            ValidateCaseCoverage(baselineScenario, candidateScenario);
            var metricImpacts = new Dictionary<string, MetricImpact>(StringComparer.Ordinal);
            foreach (var baselineCase in baselineScenario.CaseSummaries.OrderBy(item => item.Key, StringComparer.Ordinal))
            {
                var candidateCase = candidateScenario.CaseSummaries[baselineCase.Key];
                ValidateMetricCoverage(baselineScenario.Scenario, baselineCase.Key, baselineCase.Value, candidateCase);
                foreach (var baselineMetric in baselineCase.Value)
                {
                    if (baselineMetric.Value.Median == 0)
                    {
                        continue;
                    }

                    var candidateSummary = candidateCase[baselineMetric.Key];
                    var lowerIsBetter = !HigherIsBetterMetrics.Contains(baselineMetric.Key);
                    var ratio = candidateSummary.Median / baselineMetric.Value.Median;
                    var percentImprovement = lowerIsBetter ? (1 - ratio) * 100 : (ratio - 1) * 100;
                    var impactKey = $"{baselineCase.Key} / {baselineMetric.Key}";
                    metricImpacts[impactKey] = new MetricImpact(
                        baselineCase.Key,
                        baselineMetric.Key,
                        lowerIsBetter,
                        baselineMetric.Value.Median,
                        candidateSummary.Median,
                        baselineMetric.Value.P95,
                        candidateSummary.P95,
                        ratio,
                        percentImprovement,
                        percentImprovement > 0);
                }
            }

            var safetyPassed = candidateScenario.Invariants.All(invariant => invariant.Passed) &&
                candidateScenario.Reliability.Failures == 0;
            comparisons.Add(new ScenarioImpact(
                baselineScenario.PlanId,
                baselineScenario.Scenario,
                baselineScenario.Title,
                safetyPassed,
                metricImpacts));
        }

        return new BenchmarkComparisonReport(baseline.RunId, candidate.RunId, comparisons);
    }

    private static string FormatNames(string[] names) =>
        names.Length == 0 ? "none" : string.Join(", ", names.Order(StringComparer.Ordinal));

    private static bool SameExecutionEnvironment(BenchmarkEnvironment baseline, BenchmarkEnvironment candidate) =>
        string.Equals(baseline.MachineNameHash, candidate.MachineNameHash, StringComparison.Ordinal) &&
        string.Equals(baseline.ExcelVersion, candidate.ExcelVersion, StringComparison.Ordinal) &&
        string.Equals(baseline.OperatingSystem, candidate.OperatingSystem, StringComparison.Ordinal) &&
        string.Equals(baseline.ProcessArchitecture, candidate.ProcessArchitecture, StringComparison.Ordinal) &&
        baseline.LogicalProcessors == candidate.LogicalProcessors;

    private static void ValidateCaseCoverage(ScenarioResult baseline, ScenarioResult candidate)
    {
        var baselineCases = baseline.CaseSummaries.Keys.ToHashSet(StringComparer.Ordinal);
        var candidateCases = candidate.CaseSummaries.Keys.ToHashSet(StringComparer.Ordinal);
        var missing = baselineCases.Except(candidateCases, StringComparer.Ordinal).ToArray();
        var unexpected = candidateCases.Except(baselineCases, StringComparer.Ordinal).ToArray();
        if (missing.Length > 0 || unexpected.Length > 0)
        {
            throw new InvalidDataException(
                $"Scenario '{baseline.Scenario}' requires identical workload cases. " +
                $"Missing: {FormatNames(missing)}. Unexpected: {FormatNames(unexpected)}.");
        }
    }

    private static void ValidateMetricCoverage(
        string scenario,
        string caseName,
        IReadOnlyDictionary<string, DistributionSummary> baseline,
        IReadOnlyDictionary<string, DistributionSummary> candidate)
    {
        var baselineMetrics = baseline.Keys.ToHashSet(StringComparer.Ordinal);
        var candidateMetrics = candidate.Keys.ToHashSet(StringComparer.Ordinal);
        var missing = baselineMetrics.Except(candidateMetrics, StringComparer.Ordinal).ToArray();
        var unexpected = candidateMetrics.Except(baselineMetrics, StringComparer.Ordinal).ToArray();
        if (missing.Length > 0 || unexpected.Length > 0)
        {
            throw new InvalidDataException(
                $"Scenario '{scenario}', case '{caseName}' requires identical metrics. " +
                $"Missing: {FormatNames(missing)}. Unexpected: {FormatNames(unexpected)}.");
        }
    }
}
