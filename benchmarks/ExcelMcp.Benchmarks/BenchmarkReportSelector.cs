namespace Sbroenne.ExcelMcp.Benchmarks;

internal static class BenchmarkReportSelector
{
    public static BenchmarkRunReport Select(
        BenchmarkRunReport report,
        IReadOnlySet<string> requestedPlans)
    {
        ArgumentNullException.ThrowIfNull(report);
        ArgumentNullException.ThrowIfNull(requestedPlans);

        if (requestedPlans.Count == 0)
        {
            throw new ArgumentException("At least one plan must be selected.", nameof(requestedPlans));
        }

        var availablePlans = report.Scenarios
            .Select(scenario => scenario.PlanId)
            .ToHashSet(StringComparer.Ordinal);
        var missingPlans = requestedPlans
            .Except(availablePlans, StringComparer.Ordinal)
            .Order(StringComparer.Ordinal)
            .ToArray();
        if (missingPlans.Length > 0)
        {
            throw new InvalidDataException(
                $"Report '{report.RunId}' does not contain requested plan evidence: {string.Join(", ", missingPlans)}.");
        }

        var scenarios = report.Scenarios
            .Where(scenario => requestedPlans.Contains(scenario.PlanId))
            .OrderBy(scenario => scenario.PlanId, StringComparer.Ordinal)
            .ToArray();
        return report with { Scenarios = scenarios };
    }
}
