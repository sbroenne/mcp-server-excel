using Xunit;

namespace Sbroenne.ExcelMcp.Benchmarks.Tests;

[Trait("Layer", "Benchmarks")]
[Trait("Category", "Unit")]
[Trait("Feature", "Benchmarks")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class BenchmarkContractValidatorTests
{
    [Fact]
    public void ValidateScenario_MissingCatalogInvariant_RejectsIncompleteSafetyGate()
    {
        var scenario = TestReportFactory.CreateSingleScenario().Scenarios[0];

        var exception = Assert.Throws<InvalidDataException>(() => BenchmarkContractValidator.ValidateScenario(scenario));

        Assert.Contains("no_stale_read_after_write", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ValidateScenario_MissingPrimaryMetric_RejectsNonComparableRun()
    {
        var plan = BenchmarkPlanCatalog.All.Single(item => item.Id == "06");
        var scenario = TestReportFactory.CreateSingleScenario().Scenarios[0] with
        {
            Invariants = plan.ReliabilityInvariants
                .Select(name => new BenchmarkInvariant(name, true, "synthetic"))
                .ToArray()
        };

        var exception = Assert.Throws<InvalidDataException>(() => BenchmarkContractValidator.ValidateScenario(scenario));

        Assert.Contains("cold_read_ms", exception.Message, StringComparison.Ordinal);
    }
}
