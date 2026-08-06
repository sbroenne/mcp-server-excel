namespace Sbroenne.ExcelMcp.Benchmarks;

internal static class BenchmarkContractValidator
{
    public static void ValidateScenario(ScenarioResult result)
    {
        ArgumentNullException.ThrowIfNull(result);
        var plan = BenchmarkPlanCatalog.All.SingleOrDefault(item => item.Id == result.PlanId)
            ?? throw new InvalidDataException($"Scenario '{result.Scenario}' has unknown plan ID '{result.PlanId}'.");

        if (!string.Equals(plan.Scenario, result.Scenario, StringComparison.Ordinal))
        {
            throw new InvalidDataException(
                $"Plan {plan.Id} must report scenario '{plan.Scenario}', not '{result.Scenario}'.");
        }

        var invariantNames = result.Invariants.Select(item => item.Name).ToHashSet(StringComparer.Ordinal);
        var missingInvariants = plan.ReliabilityInvariants
            .Where(name => !invariantNames.Contains(name))
            .ToArray();
        if (missingInvariants.Length > 0)
        {
            throw new InvalidDataException(
                $"Scenario '{result.Scenario}' omitted required safety invariants: {string.Join(", ", missingInvariants)}.");
        }

        var metricNames = result.Observations
            .SelectMany(observation => observation.Metrics.Keys)
            .ToHashSet(StringComparer.Ordinal);
        var missingMetrics = plan.PrimaryMetrics
            .Where(name => !metricNames.Contains(name))
            .ToArray();
        if (missingMetrics.Length > 0)
        {
            throw new InvalidDataException(
                $"Scenario '{result.Scenario}' omitted required comparison metrics: {string.Join(", ", missingMetrics)}.");
        }
    }
}
