using Xunit;

namespace Sbroenne.ExcelMcp.Benchmarks.Tests;

[Trait("Layer", "Benchmarks")]
[Trait("Category", "Unit")]
[Trait("Feature", "Benchmarks")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class StatisticsTests
{
    [Fact]
    public void Summarize_KnownLatencySamples_ReturnsInterpolatedPercentilesAndDeterministicMedianInterval()
    {
        double[] samples = [10, 20, 30, 40, 50];

        var summary = Statistics.Summarize(samples, bootstrapIterations: 2_000, seed: 8675309);

        Assert.Equal(5, summary.Count);
        Assert.Equal(30, summary.Median, precision: 6);
        Assert.Equal(48, summary.P95, precision: 6);
        Assert.Equal(49.6, summary.P99, precision: 6);
        Assert.InRange(summary.MedianConfidence95.Low, 10, 30);
        Assert.InRange(summary.MedianConfidence95.High, 30, 50);
    }

    [Fact]
    public void Summarize_NoFailures_ReportsRuleOfThreeReliabilityUpperBound()
    {
        var reliability = Statistics.SummarizeReliability(successes: 100, failures: 0);

        Assert.Equal(1, reliability.SuccessRate, precision: 6);
        Assert.Equal(0, reliability.FailureRate, precision: 6);
        Assert.InRange(reliability.FailureRateConfidence95.High, 0.029, 0.031);
    }

    [Fact]
    public void Compare_LowerLatencyCandidate_ReportsDirectionAndPairedEffect()
    {
        double[] baseline = [100, 110, 90, 105];
        double[] candidate = [50, 55, 45, 52.5];

        var comparison = Statistics.ComparePaired(baseline, candidate, lowerIsBetter: true);

        Assert.True(comparison.Improved);
        Assert.Equal(0.5, comparison.CandidateToBaselineRatio, precision: 6);
        Assert.Equal(50, comparison.PercentImprovement, precision: 6);
    }
}
