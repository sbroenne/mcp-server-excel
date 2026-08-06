using System.Globalization;
using System.Text;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.Benchmarks;

internal static class ComparisonWriter
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public static ReportPaths Write(BenchmarkComparisonReport report, string outputDirectory)
    {
        Directory.CreateDirectory(outputDirectory);
        var jsonPath = Path.Combine(outputDirectory, "comparison.json");
        var markdownPath = Path.Combine(outputDirectory, "comparison.md");
        File.WriteAllText(jsonPath, JsonSerializer.Serialize(report, JsonOptions));
        File.WriteAllText(markdownPath, RenderMarkdown(report));
        return new ReportPaths(jsonPath, string.Empty, markdownPath);
    }

    private static string RenderMarkdown(BenchmarkComparisonReport report)
    {
        var builder = new StringBuilder();
        builder.AppendLine("# Excel MCP before/after comparison")
            .AppendLine()
            .AppendLine(CultureInfo.InvariantCulture, $"Baseline `{report.BaselineRunId}` compared with candidate `{report.CandidateRunId}`.")
            .AppendLine()
            .AppendLine("| Plan | Case | Metric | Baseline median | Candidate median | Improvement | Safety |")
            .AppendLine("|---:|---|---|---:|---:|---:|---|");

        foreach (var scenario in report.Scenarios)
        {
            foreach (var metric in scenario.Metrics.Values.OrderByDescending(metric => metric.PercentImprovement))
            {
                builder.Append("| ").Append(scenario.PlanId)
                    .Append(" | ").Append(metric.Case)
                    .Append(" | ").Append(metric.Metric)
                    .Append(" | ").Append(metric.BaselineMedian.ToString("0.###", CultureInfo.InvariantCulture))
                    .Append(" | ").Append(metric.CandidateMedian.ToString("0.###", CultureInfo.InvariantCulture))
                    .Append(" | ").Append(metric.PercentImprovement.ToString("+0.0;-0.0;0.0", CultureInfo.InvariantCulture)).Append('%')
                    .Append(" | ").Append(scenario.SafetyInvariantsPassed ? "PASS" : "FAIL")
                    .AppendLine(" |");
            }
        }

        builder.AppendLine()
            .AppendLine("A faster result is not accepted when the candidate fails a reliability invariant.");
        return builder.ToString();
    }
}
