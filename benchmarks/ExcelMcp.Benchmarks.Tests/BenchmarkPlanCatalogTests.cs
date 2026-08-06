using Xunit;

namespace Sbroenne.ExcelMcp.Benchmarks.Tests;

[Trait("Layer", "Benchmarks")]
[Trait("Category", "Unit")]
[Trait("Feature", "Benchmarks")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class BenchmarkPlanCatalogTests
{
    [Fact]
    public void Plans_FirstNineImprovementIdeas_HaveUniqueComparableMeasurementContracts()
    {
        var plans = BenchmarkPlanCatalog.All;

        Assert.Equal(9, plans.Count);
        Assert.Equal(
            ["01", "02", "03", "04", "05", "06", "07", "08", "09"],
            plans.Select(plan => plan.Id));
        Assert.Equal(plans.Count, plans.Select(plan => plan.Scenario).Distinct(StringComparer.Ordinal).Count());

        Assert.All(plans, plan =>
        {
            Assert.NotEmpty(plan.PrimaryMetrics);
            Assert.NotEmpty(plan.ReliabilityInvariants);
            Assert.False(string.IsNullOrWhiteSpace(plan.BaselineMeaning));
            Assert.False(string.IsNullOrWhiteSpace(plan.CandidateSuccess));
        });

        Assert.Contains(plans, plan => plan.PrimaryMetrics.Contains("token_estimate", StringComparer.Ordinal));
        Assert.Contains(plans, plan => plan.PrimaryMetrics.Contains("refresh_to_consistent_read_ms", StringComparer.Ordinal));
    }
}
