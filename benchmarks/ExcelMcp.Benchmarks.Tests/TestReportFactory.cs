namespace Sbroenne.ExcelMcp.Benchmarks.Tests;

internal static class TestReportFactory
{
    public static BenchmarkRunReport CreateSingleScenario()
    {
        var observations = new List<BenchmarkObservation>
        {
            new(0, "warm", true, null, new Dictionary<string, double> { ["warm_read_ms"] = 20 }),
            new(1, "warm", true, null, new Dictionary<string, double> { ["warm_read_ms"] = 30 }),
            new(2, "warm", true, null, new Dictionary<string, double> { ["warm_read_ms"] = 40 })
        };

        var scenario = ScenarioResult.Create(
            "06",
            "read-fast-path",
            "Read fast path",
            observations,
            [new BenchmarkInvariant("round_trip_values_equal", true, "Known literal matched")],
            "synthetic test workload");

        return new BenchmarkRunReport(
            "test-run",
            "test",
            new DateTimeOffset(2026, 8, 5, 12, 0, 0, TimeSpan.Zero),
            new DateTimeOffset(2026, 8, 5, 12, 0, 1, TimeSpan.Zero),
            new BenchmarkConfiguration(0, 3, 3, false),
            new BenchmarkEnvironment("test", ".NET test", "Windows test", "Excel test", "x64", 8, 16_000_000_000, "abc123", "test", false),
            [scenario]);
    }
}
