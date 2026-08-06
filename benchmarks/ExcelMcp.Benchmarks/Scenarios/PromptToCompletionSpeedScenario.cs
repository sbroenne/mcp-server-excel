using System.Text;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.Benchmarks.Scenarios;

/// <summary>
/// Paired, same-run prompt-to-completion experiment using the public MCP transport.
/// Each case has a fresh client so initialize, tools/list, and tools/call bytes are comparable.
/// </summary>
internal sealed class PromptToCompletionSpeedScenario : IBenchmarkScenario
{
    private const int OperationCount = 8;

    internal static IReadOnlyList<string> Cases { get; } =
    [
        "prompt-to-completion-legacy",
        "prompt-to-completion-execute-plan",
        "prompt-to-completion-execute-plan-open-and-describe"
    ];

    public string PlanId => "10";

    public string Name => "prompt-to-completion-speed";

    public async Task<ScenarioResult> RunAsync(BenchmarkContext context, CancellationToken cancellationToken)
    {
        var masterPath = context.CreateWorkingPath("prompt-speed-master");
        BenchmarkContext.CreateDataWorkbook(masterPath, rows: OperationCount + 2, columns: 2);

        var observations = new List<BenchmarkObservation>();
        var exactValues = true;
        var noLostOrDuplicateOperations = true;
        var validCompactSummaries = true;
        var noUnknownOutcomes = true;
        var sessionsClosed = true;
        var mcpTransportSucceeded = true;
        var total = context.Configuration.Warmups + context.Configuration.Iterations;
        var variants = new[]
        {
            PromptWorkflowVariant.Legacy,
            PromptWorkflowVariant.ExecutePlanOnly,
            PromptWorkflowVariant.ExecutePlanAndOpenDescribe
        };

        for (var iteration = 0; iteration < total; iteration++)
        {
            // Rotate each three-case cohort to balance Excel startup and thermal drift.
            foreach (var variant in variants.Skip(iteration % variants.Length).Concat(variants.Take(iteration % variants.Length)))
            {
                cancellationToken.ThrowIfCancellationRequested();
                var result = await RunCaseAsync(context, masterPath, iteration, variant, cancellationToken);
                exactValues &= result.ExactValues;
                noLostOrDuplicateOperations &= result.NoLostOrDuplicateOperations;
                validCompactSummaries &= result.ValidCompactSummary;
                noUnknownOutcomes &= result.Run.KnownOutcome;
                sessionsClosed &= result.Run.SessionClosed;
                mcpTransportSucceeded &= result.Run.Success;

                if (iteration >= context.Configuration.Warmups)
                {
                    observations.Add(result.ToObservation(iteration - context.Configuration.Warmups));
                }
            }
        }

        var plan = BenchmarkPlanCatalog.All.Single(plan => plan.Id == PlanId);
        return ScenarioResult.Create(
            PlanId,
            Name,
            plan.Title,
            observations,
            [
                new BenchmarkInvariant("exact_values", exactValues, $"All {OperationCount} verified cells matched their ordered literal writes: {exactValues}"),
                new BenchmarkInvariant("no_lost_or_duplicate_operations", noLostOrDuplicateOperations, $"Every verified write appeared once and only once: {noLostOrDuplicateOperations}"),
                new BenchmarkInvariant("session_cleanup", sessionsClosed, $"Every public MCP case closed its owned session with save:false: {sessionsClosed}"),
                new BenchmarkInvariant("valid_compact_summary", validCompactSummaries, $"Every case generated a non-empty compact completion summary: {validCompactSummaries}"),
                new BenchmarkInvariant("no_unknown_outcome", noUnknownOutcomes, $"Every write workflow returned a known completion or failure outcome: {noUnknownOutcomes}"),
                new BenchmarkInvariant("mcp_transport", mcpTransportSucceeded, $"Every public MCP workflow completed successfully: {mcpTransportSucceeded}")
            ],
            "Three explicit public-MCP workflows under the Copilot compact tool profile: legacy calls, execute-plan-only, and execute-plan plus open-and-describe. Optimized cases use the plan's bounded final verification receipt instead of a second measured read-back call.",
            [
                "The fixture workbook is created before timing. Each case operates on an independent copy and closes through file(close, save:false) in a finally block.",
                "The prompt workflow starts the server with --tool-profile copilot-compact; full remains the default for other clients.",
                "For optimized cases, an independent range.get-values correctness audit runs on the live session but is excluded from prompt latency, request, byte, and token metrics. This keeps exact-value invariants independent while measuring the agent-visible one-call verification workflow.",
                "Token counts are deterministic ceil(UTF-8 wire bytes / 4) estimates. Byte metrics retain the exact client-to-server and server-to-client breakdown for initialization, tools/list, and tools/call."
            ]);
    }

    private static async Task<CaseResult> RunCaseAsync(
        BenchmarkContext context,
        string masterPath,
        int iteration,
        PromptWorkflowVariant variant,
        CancellationToken cancellationToken)
    {
        var workbookPath = context.CopyWorkbook(masterPath, CaseName(variant));
        var writes = CreateWrites(iteration);
        var run = await ProtocolFootprintProbe.RunPromptWorkflowAsync(
            variant,
            workbookPath,
            writes,
            context.Configuration.ShowExcel,
            cancellationToken);
        var expectedValues = writes.Select(write => write.Value).ToArray();
        var exactValues = run.Success && run.Values.SequenceEqual(expectedValues);
        var noLostOrDuplicate = exactValues && run.Values.Distinct().Count() == OperationCount;
        var summary = BuildCompletionSummary(run.Description, run.Values);

        return new CaseResult(
            CaseName(variant),
            variant,
            run,
            exactValues,
            noLostOrDuplicate,
            run.Success && IsValidSummary(summary),
            Encoding.UTF8.GetByteCount(summary));
    }

    private static PromptWorkflowWrite[] CreateWrites(int iteration) =>
        Enumerable.Range(0, OperationCount)
            .Select(index => new PromptWorkflowWrite($"A{index + 2}", iteration * 10_000d + index + 1))
            .ToArray();

    private static string CaseName(PromptWorkflowVariant variant) => variant switch
    {
        PromptWorkflowVariant.Legacy => Cases[0],
        PromptWorkflowVariant.ExecutePlanOnly => Cases[1],
        PromptWorkflowVariant.ExecutePlanAndOpenDescribe => Cases[2],
        _ => throw new ArgumentOutOfRangeException(nameof(variant))
    };

    private static string BuildCompletionSummary(string description, IReadOnlyList<double> verifiedValues) =>
        JsonSerializer.Serialize(new
        {
            description = JsonSerializer.Deserialize<JsonElement>(description),
            verifiedCount = verifiedValues.Count,
            firstValue = verifiedValues.Count > 0 ? verifiedValues[0] : 0,
            lastValue = verifiedValues.Count > 0 ? verifiedValues[^1] : 0
        });

    private static bool IsValidSummary(string summary)
    {
        try
        {
            using var document = JsonDocument.Parse(summary);
            return document.RootElement.ValueKind == JsonValueKind.Object && document.RootElement.EnumerateObject().Any();
        }
        catch (JsonException)
        {
            return false;
        }
    }

    private sealed record CaseResult(
        string Case,
        PromptWorkflowVariant Variant,
        PromptWorkflowRunResult Run,
        bool ExactValues,
        bool NoLostOrDuplicateOperations,
        bool ValidCompactSummary,
        long SummaryPayloadBytes)
    {
        public BenchmarkObservation ToObservation(int iteration) => new(
            iteration,
            Case,
            ExactValues && NoLostOrDuplicateOperations && ValidCompactSummary && Run.KnownOutcome && Run.Success,
            Run.Error,
            new Dictionary<string, double>
            {
                ["prompt_to_completion_ms"] = Run.PromptToCompletionMilliseconds,
                ["open_describe_ms"] = Run.OpenDescribeMilliseconds,
                ["execution_ms"] = Run.ExecutionMilliseconds,
                ["verification_ms"] = Run.VerificationMilliseconds,
                ["request_count"] = Run.ToolCallCount,
                ["payload_bytes"] = Run.WireBytes.TotalBytes,
                ["token_estimate"] = BenchmarkContext.EstimateTokensFromUtf8Bytes(Run.WireBytes.TotalBytes),
                ["mcp_initialize_request_bytes"] = Run.WireBytes.InitializeRequestBytes,
                ["mcp_initialize_response_bytes"] = Run.WireBytes.InitializeResponseBytes,
                ["mcp_tools_list_request_bytes"] = Run.WireBytes.ToolsListRequestBytes,
                ["mcp_tools_list_response_bytes"] = Run.WireBytes.ToolsListResponseBytes,
                ["mcp_tool_call_request_bytes"] = Run.WireBytes.ToolCallRequestBytes,
                ["mcp_tool_call_response_bytes"] = Run.WireBytes.ToolCallResponseBytes,
                ["summary_payload_bytes"] = SummaryPayloadBytes,
                ["operations_per_second"] = Run.ExecutionMilliseconds > 0 ? OperationCount / (Run.ExecutionMilliseconds / 1000d) : 0
            },
            new Dictionary<string, string>
            {
                ["mode"] = Variant.ToString(),
                ["implementation"] = Variant switch
                {
                    PromptWorkflowVariant.Legacy => "public-mcp-separate-calls",
                    PromptWorkflowVariant.ExecutePlanOnly => "public-mcp-workflow.execute-plan",
                    PromptWorkflowVariant.ExecutePlanAndOpenDescribe => "public-mcp-workflow.open-and-describe+execute-plan",
                    _ => throw new InvalidOperationException($"Unknown prompt workflow variant: {Variant}")
                },
                ["operation_count"] = OperationCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["token_measurement"] = "ceil(utf8-wire-bytes/4); deterministic estimate, not model-specific"
            },
            Run.KnownOutcome ? (Run.Success ? "completed" : "failed") : "outcome-unknown");
    }
}
