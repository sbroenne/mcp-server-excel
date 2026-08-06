using Xunit;

namespace Sbroenne.ExcelMcp.Benchmarks.Tests;

[Trait("Layer", "Benchmarks")]
[Trait("Category", "Unit")]
[Trait("Feature", "Benchmarks")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ComparisonEngineTests
{
    [Fact]
    public void Compare_SameScenarioAndMetric_ComputesBaselineRelativeImpact()
    {
        var baseline = TestReportFactory.CreateSingleScenario();
        var candidateObservations = baseline.Scenarios[0].Observations
            .Select(observation => observation with
            {
                Metrics = new Dictionary<string, double>
                {
                    ["warm_read_ms"] = observation.Metrics["warm_read_ms"] / 2
                }
            })
            .ToArray();
        var candidateScenario = ScenarioResult.Create(
            "06",
            "read-fast-path",
            "Read fast path",
            candidateObservations,
            baseline.Scenarios[0].Invariants,
            "synthetic candidate");
        var candidate = baseline with { RunId = "candidate", Scenarios = [candidateScenario] };

        var comparison = ComparisonEngine.Compare(baseline, candidate);

        var metric = Assert.Single(comparison.Scenarios).Metrics["warm / warm_read_ms"];
        Assert.Equal("warm", metric.Case);
        Assert.Equal(50, metric.PercentImprovement, precision: 6);
        Assert.True(metric.Improved);
        Assert.True(comparison.Scenarios[0].SafetyInvariantsPassed);
    }

    [Fact]
    public void Compare_CandidateOmitsBaselineScenario_RejectsIncompleteComparison()
    {
        var baseline = TestReportFactory.CreateSingleScenario();
        var candidate = baseline with { RunId = "candidate", Scenarios = [] };

        var exception = Assert.Throws<InvalidDataException>(() => ComparisonEngine.Compare(baseline, candidate));

        Assert.Contains("read-fast-path", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Compare_DifferentRepetitionConfiguration_RejectsUnlikeExperiment()
    {
        var baseline = TestReportFactory.CreateSingleScenario();
        var candidate = baseline with
        {
            RunId = "candidate",
            Configuration = baseline.Configuration with { Iterations = baseline.Configuration.Iterations + 1 }
        };

        var exception = Assert.Throws<InvalidDataException>(() => ComparisonEngine.Compare(baseline, candidate));

        Assert.Contains("configuration", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Compare_InProgressReport_RejectsIncompleteRun()
    {
        var baseline = TestReportFactory.CreateSingleScenario() with { RunState = "in-progress" };
        var candidate = TestReportFactory.CreateSingleScenario() with { RunId = "candidate" };

        var exception = Assert.Throws<InvalidDataException>(() => ComparisonEngine.Compare(baseline, candidate));

        Assert.Contains("completed", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}
