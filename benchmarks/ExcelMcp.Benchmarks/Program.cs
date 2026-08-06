using System.Text.Json;

namespace Sbroenne.ExcelMcp.Benchmarks;

internal static class Program
{
    private static readonly JsonSerializerOptions CatalogJsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public static async Task<int> Main(string[] args)
    {
        Environment.SetEnvironmentVariable("EXCELMCP_TELEMETRY_OPTOUT", "true");

        try
        {
            var options = BenchmarkOptions.Parse(args);
            if (options.Command == BenchmarkCommand.Catalog)
            {
                Console.WriteLine(JsonSerializer.Serialize(BenchmarkPlanCatalog.All, CatalogJsonOptions));
                return 0;
            }

            if (options.Command == BenchmarkCommand.Compare)
            {
                var baseline = ReportWriter.Read(options.BaselinePath!);
                var candidate = ReportWriter.Read(options.CandidatePath!);
                if (options.SelectedPlans.Count < BenchmarkPlanCatalog.All.Count)
                {
                    baseline = BenchmarkReportSelector.Select(baseline, options.SelectedPlans);
                    candidate = BenchmarkReportSelector.Select(candidate, options.SelectedPlans);
                }

                var comparison = ComparisonEngine.Compare(baseline, candidate);
                var paths = ComparisonWriter.Write(comparison, options.OutputDirectory);
                Console.WriteLine($"Comparison JSON: {paths.JsonPath}");
                Console.WriteLine($"Comparison Markdown: {paths.MarkdownPath}");
                return comparison.Scenarios.All(scenario => scenario.SafetyInvariantsPassed) ? 0 : 2;
            }

            using var cancellationSource = new CancellationTokenSource(options.MaximumRunTime);
            var runner = new BenchmarkRunner(options);
            var report = await runner.RunAsync(cancellationSource.Token);
            var reportPaths = ReportWriter.Write(report, options.OutputDirectory);
            Console.WriteLine($"Baseline JSON: {reportPaths.JsonPath}");
            Console.WriteLine($"Raw observations: {reportPaths.CsvPath}");
            Console.WriteLine($"Readable report: {reportPaths.MarkdownPath}");
            return report.Scenarios.All(scenario => scenario.Invariants.All(invariant => invariant.Passed)) ? 0 : 2;
        }
        catch (OperationCanceledException)
        {
            Console.Error.WriteLine("Benchmark run exceeded its configured maximum duration or was cancelled.");
            return 3;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine($"Benchmark failed: {exception.GetType().Name}: {exception.Message}");
            Console.Error.WriteLine(exception.StackTrace);
            return 1;
        }
    }
}
