using Xunit;

namespace Sbroenne.ExcelMcp.Benchmarks.Tests;

[Trait("Layer", "Benchmarks")]
[Trait("Category", "Unit")]
[Trait("Feature", "Benchmarks")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ScenarioResultTests
{
    [Fact]
    public void Create_FailedSafetyObservation_StillContributesPerformanceMeasurement()
    {
        var observations = new BenchmarkObservation[]
        {
            new(0, "small", true, null, new Dictionary<string, double> { ["latency_ms"] = 10 }),
            new(1, "large", false, "safety gate failed", new Dictionary<string, double> { ["latency_ms"] = 110 })
        };

        var result = ScenarioResult.Create(
            "test",
            "synthetic",
            "Synthetic",
            observations,
            [],
            "synthetic workload");

        Assert.Equal(2, result.Summaries["latency_ms"].Count);
        Assert.Equal(60, result.Summaries["latency_ms"].Median);
    }

    [Fact]
    public void Create_DifferentWorkloadCases_ProducesSeparateDistributions()
    {
        var observations = new BenchmarkObservation[]
        {
            new(0, "small", true, null, new Dictionary<string, double> { ["latency_ms"] = 10 }),
            new(0, "large", true, null, new Dictionary<string, double> { ["latency_ms"] = 110 })
        };

        var result = ScenarioResult.Create(
            "test",
            "synthetic",
            "Synthetic",
            observations,
            [],
            "synthetic workload");

        Assert.Equal(10, result.CaseSummaries["small"]["latency_ms"].Median);
        Assert.Equal(110, result.CaseSummaries["large"]["latency_ms"].Median);
    }
}
