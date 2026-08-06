namespace Sbroenne.ExcelMcp.Benchmarks;

internal interface IBenchmarkScenario
{
    string PlanId { get; }

    string Name { get; }

    Task<ScenarioResult> RunAsync(BenchmarkContext context, CancellationToken cancellationToken);
}
