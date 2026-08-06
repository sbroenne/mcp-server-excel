using Xunit;

namespace Sbroenne.ExcelMcp.Benchmarks.Tests;

[Trait("Layer", "Benchmarks")]
[Trait("Category", "Unit")]
[Trait("Feature", "Benchmarks")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class BenchmarkReportSelectorTests
{
    [Fact]
    public void Select_RequestedPlans_PreservesOnlyEquivalentScenarioEvidence()
    {
        var source = TestReportFactory.CreateSingleScenario();
        var plan05 = source.Scenarios[0] with
        {
            PlanId = "05",
            Scenario = "vectorized-writes",
            Title = "Vectorized writes"
        };
        var report = source with { Scenarios = [plan05, source.Scenarios[0]] };

        var selected = BenchmarkReportSelector.Select(report, new HashSet<string>(["05"], StringComparer.Ordinal));

        var scenario = Assert.Single(selected.Scenarios);
        Assert.Equal("05", scenario.PlanId);
        Assert.Equal(report.RunId, selected.RunId);
        Assert.Equal(report.Configuration, selected.Configuration);
        Assert.Equal(report.Environment, selected.Environment);
    }

    [Fact]
    public void Select_MissingRequestedPlan_RejectsPartialEvidence()
    {
        var report = TestReportFactory.CreateSingleScenario();

        var exception = Assert.Throws<InvalidDataException>(() =>
            BenchmarkReportSelector.Select(report, new HashSet<string>(["05"], StringComparer.Ordinal)));

        Assert.Contains("05", exception.Message, StringComparison.Ordinal);
    }
}
