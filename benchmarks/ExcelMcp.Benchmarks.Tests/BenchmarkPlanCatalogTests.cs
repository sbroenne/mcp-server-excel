using Xunit;
using Sbroenne.ExcelMcp.Benchmarks.Scenarios;

namespace Sbroenne.ExcelMcp.Benchmarks.Tests;

[Trait("Layer", "Benchmarks")]
[Trait("Category", "Unit")]
[Trait("Feature", "Benchmarks")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class BenchmarkPlanCatalogTests
{
    [Fact]
    public void Plans_ImprovementIdeas_HaveUniqueComparableMeasurementContracts()
    {
        var plans = BenchmarkPlanCatalog.All;

        Assert.Equal(10, plans.Count);
        Assert.Equal(
            ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10"],
            plans.Select(plan => plan.Id));
        Assert.Equal(plans.Count, plans.Select(plan => plan.Scenario).Distinct(StringComparer.Ordinal).Count());

        Assert.All(plans, plan =>
        {
            Assert.NotEmpty(plan.PrimaryMetrics);
            Assert.NotEmpty(plan.ReliabilityInvariants);
            Assert.False(string.IsNullOrWhiteSpace(plan.BaselineMeaning));
            Assert.False(string.IsNullOrWhiteSpace(plan.CandidateSuccess));
        });

        Assert.Contains(plans, plan => plan.PrimaryMetrics.Contains("token_estimate", StringComparer.Ordinal));
        Assert.Contains(plans, plan => plan.PrimaryMetrics.Contains("refresh_to_consistent_read_ms", StringComparer.Ordinal));
    }

    [Fact]
    public void Plan10_PromptToCompletionSpeed_HasTheCumulativeWorkflowContract()
    {
        var plan = BenchmarkPlanCatalog.All.Single(item => item.Id == "10");

        Assert.Equal("prompt-to-completion-speed", plan.Scenario);
        Assert.Equal(
            [
                "prompt_to_completion_ms",
                "open_describe_ms",
                "execution_ms",
                "verification_ms",
                "request_count",
                "payload_bytes",
                "token_estimate",
                "mcp_initialize_request_bytes",
                "mcp_initialize_response_bytes",
                "mcp_tools_list_request_bytes",
                "mcp_tools_list_response_bytes",
                "mcp_tool_call_request_bytes",
                "mcp_tool_call_response_bytes",
                "summary_payload_bytes",
                "operations_per_second"
            ],
            plan.PrimaryMetrics);
        Assert.Equal(
            ["exact_values", "no_lost_or_duplicate_operations", "session_cleanup", "valid_compact_summary", "no_unknown_outcome", "mcp_transport"],
            plan.ReliabilityInvariants);
    }

    [Fact]
    public void Plan10_PromptToCompletionSpeed_ReportsPairedLegacyAndCandidateCases()
    {
        Assert.Equal(
            [
                "prompt-to-completion-legacy",
                "prompt-to-completion-execute-plan",
                "prompt-to-completion-execute-plan-open-and-describe"
            ],
            PromptToCompletionSpeedScenario.Cases);
    }
}
