using Sbroenne.ExcelMcp.Benchmarks.Scenarios;

namespace Sbroenne.ExcelMcp.Benchmarks;

internal sealed class BenchmarkRunner(BenchmarkOptions options)
{
    private readonly IReadOnlyList<IBenchmarkScenario> _scenarios =
    [
        new TimeoutQuarantineScenario(),
        new BoundedWorkbookQueueScenario(),
        new TargetedSafetyInspectionScenario(),
        new ServerSideBatchScenario(),
        new VectorizedWritesScenario(),
        new ReadFastPathScenario(),
        new IdempotencyScenario(),
        new DurableJournalCheckpointScenario(),
        new PreciseProcessTrackingScenario(),
        new PromptToCompletionSpeedScenario()
    ];

    public async Task<BenchmarkRunReport> RunAsync(CancellationToken cancellationToken)
    {
        var startedAt = DateTimeOffset.UtcNow;
        using var context = new BenchmarkContext(options);
        Console.WriteLine("Capturing Excel, .NET, OS, and Git environment metadata...");
        var environment = EnvironmentProbe.Capture(context);
        var runId = $"{startedAt:yyyyMMdd-HHmmss}-{ShortCommit(environment.GitCommit)}";
        var results = new List<ScenarioResult>();

        foreach (var scenario in _scenarios.Where(item => options.SelectedPlans.Contains(item.PlanId)))
        {
            cancellationToken.ThrowIfCancellationRequested();
            Console.WriteLine($"[{scenario.PlanId}/{BenchmarkPlanCatalog.All.Count:D2}] {scenario.Name}...");
            try
            {
                var result = await scenario.RunAsync(context, cancellationToken);
                BenchmarkContractValidator.ValidateScenario(result);
                results.Add(result);
                Console.WriteLine($"[{scenario.PlanId}/{BenchmarkPlanCatalog.All.Count:D2}] {result.Status}: {result.Observations.Count} observations");
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception exception)
            {
                var plan = BenchmarkPlanCatalog.All.Single(plan => plan.Id == scenario.PlanId);
                results.Add(ScenarioResult.Create(
                    scenario.PlanId,
                    scenario.Name,
                    plan.Title,
                    [new BenchmarkObservation(0, "scenario-error", false, $"{exception.GetType().Name}: {exception.Message}", new Dictionary<string, double>())],
                    [new BenchmarkInvariant("scenario_completed", false, $"{exception.GetType().Name}: {exception.Message}")],
                    "Scenario setup or execution failed before a complete baseline could be captured.",
                    [exception.StackTrace ?? "No stack trace was available."],
                    status: "error"));
                Console.WriteLine($"[{scenario.PlanId}/{BenchmarkPlanCatalog.All.Count:D2}] ERROR: {exception.GetType().Name}: {exception.Message}");
            }

            WriteProgressReport(runId, startedAt, environment, results);
        }

        return new BenchmarkRunReport(
            runId,
            options.Profile,
            startedAt,
            DateTimeOffset.UtcNow,
            options.Configuration,
            environment,
            results);
    }

    private void WriteProgressReport(
        string runId,
        DateTimeOffset startedAt,
        BenchmarkEnvironment environment,
        IReadOnlyList<ScenarioResult> results)
    {
        var progress = new BenchmarkRunReport(
            runId,
            options.Profile,
            startedAt,
            DateTimeOffset.UtcNow,
            options.Configuration,
            environment,
            results,
            RunState: "in-progress");
        _ = ReportWriter.Write(progress, options.OutputDirectory);
    }

    private static string ShortCommit(string commit) => commit.Length >= 8 ? commit[..8] : commit;
}
