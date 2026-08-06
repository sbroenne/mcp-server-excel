using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Benchmarks;

internal static class ReportWriter
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    public static ReportPaths Write(BenchmarkRunReport report, string outputDirectory)
    {
        ArgumentNullException.ThrowIfNull(report);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputDirectory);

        Directory.CreateDirectory(outputDirectory);
        var jsonPath = Path.Combine(outputDirectory, "baseline.json");
        var csvPath = Path.Combine(outputDirectory, "observations.csv");
        var markdownPath = Path.Combine(outputDirectory, "baseline.md");

        File.WriteAllText(jsonPath, JsonSerializer.Serialize(report, JsonOptions));
        File.WriteAllText(csvPath, RenderCsv(report));
        File.WriteAllText(markdownPath, RenderMarkdown(report));

        return new ReportPaths(jsonPath, csvPath, markdownPath);
    }

    public static BenchmarkRunReport Read(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        return JsonSerializer.Deserialize<BenchmarkRunReport>(File.ReadAllText(path), JsonOptions)
            ?? throw new InvalidDataException($"Benchmark report '{path}' was empty or invalid.");
    }

    private static string RenderCsv(BenchmarkRunReport report)
    {
        var builder = new StringBuilder();
        builder.AppendLine("run_id,plan_id,scenario,case,iteration,success,outcome,metric,value,error");
        foreach (var scenario in report.Scenarios)
        {
            foreach (var observation in scenario.Observations)
            {
                if (observation.Metrics.Count == 0)
                {
                    AppendCsvRow(builder, report.RunId, scenario, observation, string.Empty, null);
                    continue;
                }

                foreach (var metric in observation.Metrics.OrderBy(metric => metric.Key, StringComparer.Ordinal))
                {
                    AppendCsvRow(builder, report.RunId, scenario, observation, metric.Key, metric.Value);
                }
            }
        }

        return builder.ToString();
    }

    private static void AppendCsvRow(
        StringBuilder builder,
        string runId,
        ScenarioResult scenario,
        BenchmarkObservation observation,
        string metric,
        double? value)
    {
        builder.Append(Csv(runId)).Append(',')
            .Append(Csv(scenario.PlanId)).Append(',')
            .Append(Csv(scenario.Scenario)).Append(',')
            .Append(Csv(observation.Case)).Append(',')
            .Append(observation.Iteration.ToString(CultureInfo.InvariantCulture)).Append(',')
            .Append(observation.Success ? "true" : "false").Append(',')
            .Append(Csv(observation.Outcome ?? string.Empty)).Append(',')
            .Append(Csv(metric)).Append(',')
            .Append(value?.ToString("G17", CultureInfo.InvariantCulture) ?? string.Empty).Append(',')
            .Append(Csv(observation.Error ?? string.Empty))
            .AppendLine();
    }

    // Markdown is an invariant machine artifact: numbers use Format(..., InvariantCulture)
    // and timestamps use the round-trip "O" format below.
#pragma warning disable CA1305
    private static string RenderMarkdown(BenchmarkRunReport report)
    {
        var builder = new StringBuilder();
        builder.AppendLine("# Excel MCP baseline benchmark")
            .AppendLine()
            .AppendLine($"Run `{report.RunId}` used profile `{report.Profile}` from {report.StartedAtUtc:O} to {report.FinishedAtUtc:O}. State: `{report.RunState}`.")
            .AppendLine()
            .AppendLine($"Configuration: {report.Configuration.Warmups} warmups, {report.Configuration.Iterations} measured iterations, {report.Configuration.ReliabilityIterations} reliability iterations, Excel visible: {report.Configuration.ShowExcel}.")
            .AppendLine()
            .AppendLine("## Environment")
            .AppendLine()
            .AppendLine($"- Excel: {report.Environment.ExcelVersion}")
            .AppendLine($"- .NET: {report.Environment.DotNetVersion}")
            .AppendLine($"- OS: {report.Environment.OperatingSystem}")
            .AppendLine($"- Architecture: {report.Environment.ProcessArchitecture}")
            .AppendLine($"- Git: `{report.Environment.GitBranch}` at `{report.Environment.GitCommit}`, dirty: {report.Environment.GitDirty}")
            .AppendLine()
            .AppendLine("## Scenario overview")
            .AppendLine()
            .AppendLine("| Plan | Scenario | Status | Successes | Failures |")
            .AppendLine("|---:|---|---|---:|---:|");

        foreach (var scenario in report.Scenarios.OrderBy(scenario => scenario.PlanId, StringComparer.Ordinal))
        {
            builder.AppendLine($"| {scenario.PlanId} | {EscapePipe(scenario.Title)} | {scenario.Status} | {scenario.Reliability.Successes} | {scenario.Reliability.Failures} |");
        }

        foreach (var scenario in report.Scenarios.OrderBy(scenario => scenario.PlanId, StringComparer.Ordinal))
        {
            builder.AppendLine()
                .AppendLine($"## {scenario.PlanId}. {scenario.Title}")
                .AppendLine()
                .AppendLine(scenario.Workload)
                .AppendLine()
                .AppendLine("| Metric | n | Median | p95 | p99 | 95% median CI |")
                .AppendLine("|---|---:|---:|---:|---:|---:|");

            foreach (var metric in scenario.Summaries)
            {
                var summary = metric.Value;
                builder.AppendLine(
                    $"| {metric.Key} | {summary.Count} | {Format(summary.Median)} | {Format(summary.P95)} | {Format(summary.P99)} | {Format(summary.MedianConfidence95.Low)} to {Format(summary.MedianConfidence95.High)} |");
            }

            builder.AppendLine()
                .AppendLine("Case-level distributions:")
                .AppendLine()
                .AppendLine("| Case | Metric | n | Median | p95 | p99 | 95% median CI |")
                .AppendLine("|---|---|---:|---:|---:|---:|---:|");
            foreach (var caseSummary in scenario.CaseSummaries)
            {
                foreach (var metric in caseSummary.Value)
                {
                    var summary = metric.Value;
                    builder.AppendLine(
                        $"| {EscapePipe(caseSummary.Key)} | {metric.Key} | {summary.Count} | {Format(summary.Median)} | {Format(summary.P95)} | {Format(summary.P99)} | {Format(summary.MedianConfidence95.Low)} to {Format(summary.MedianConfidence95.High)} |");
                }
            }

            builder.AppendLine()
                .AppendLine("Reliability invariants:")
                .AppendLine();
            foreach (var invariant in scenario.Invariants)
            {
                builder.AppendLine($"- {(invariant.Passed ? "PASS" : "FAIL")} `{invariant.Name}` - {invariant.Evidence}");
            }

            if (scenario.Notes.Count > 0)
            {
                builder.AppendLine()
                    .AppendLine("Notes:")
                    .AppendLine();
                foreach (var note in scenario.Notes)
                {
                    builder.AppendLine($"- {note}");
                }
            }
        }

        builder.AppendLine()
            .AppendLine("## Interpretation rules")
            .AppendLine()
            .AppendLine("- Compare medians and p95/p99 tails; do not rank changes by mean alone.")
            .AppendLine("- A performance gain is acceptable only when every hard safety invariant still passes.")
            .AppendLine("- Token figures labeled as estimates are deterministic UTF-8 payload proxies, not model-specific tokenizer counts.")
            .AppendLine("- Zero observed failures gives only the reported rule-of-three upper bound; larger reliability samples tighten it.");

        return builder.ToString();
    }
#pragma warning restore CA1305

    private static string Csv(string value) => $"\"{value.Replace("\"", "\"\"", StringComparison.Ordinal)}\"";

    private static string EscapePipe(string value) => value.Replace("|", "\\|", StringComparison.Ordinal);

    private static string Format(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);
}
