using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Benchmarks.Tests;

[Trait("Layer", "Benchmarks")]
[Trait("Category", "Unit")]
[Trait("Feature", "Benchmarks")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ReportWriterTests
{
    [Fact]
    public void Write_BenchmarkRun_PreservesRawSamplesAndProducesReadableSummary()
    {
        var outputDirectory = Path.Combine(Path.GetTempPath(), $"ExcelMcpBenchmarkReport-{Guid.NewGuid():N}");
        try
        {
            var report = TestReportFactory.CreateSingleScenario();

            var paths = ReportWriter.Write(report, outputDirectory);

            Assert.True(File.Exists(paths.JsonPath));
            Assert.True(File.Exists(paths.CsvPath));
            Assert.True(File.Exists(paths.MarkdownPath));

            using var json = JsonDocument.Parse(File.ReadAllText(paths.JsonPath));
            Assert.Equal("completed", json.RootElement.GetProperty("runState").GetString());
            var scenario = json.RootElement.GetProperty("scenarios")[0];
            Assert.Equal(3, scenario.GetProperty("observations").GetArrayLength());
            Assert.Equal(30, scenario.GetProperty("summaries").GetProperty("warm_read_ms").GetProperty("median").GetDouble());

            var markdown = File.ReadAllText(paths.MarkdownPath);
            Assert.Contains("Read fast path", markdown, StringComparison.Ordinal);
            Assert.Contains("95% median CI", markdown, StringComparison.Ordinal);
            Assert.Contains("round_trip_values_equal", markdown, StringComparison.Ordinal);
        }
        finally
        {
            if (Directory.Exists(outputDirectory))
            {
                Directory.Delete(outputDirectory, recursive: true);
            }
        }
    }
}
