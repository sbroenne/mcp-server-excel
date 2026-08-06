using System.Diagnostics;
using System.Text;
using System.Text.Json;
using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.Benchmarks.Scenarios;

internal sealed class IdempotencyScenario : IBenchmarkScenario
{
    private const int ColumnCount = 10;

    public string PlanId => "07";

    public string Name => "idempotency-keys";

    public async Task<ScenarioResult> RunAsync(BenchmarkContext context, CancellationToken cancellationToken)
    {
        var masterPath = context.CreateWorkingPath("idempotency-master");
        BenchmarkContext.CreateDataWorkbook(masterPath, rows: 2, columns: ColumnCount, includeTable: true);
        var workbookPath = context.CopyWorkbook(masterPath, "idempotency");
        var safetyRoot = context.CreateSafetyRoot("idempotency");
        var observations = new List<BenchmarkObservation>();
        var executeOnceEveryTime = true;
        var changedArgumentsNeverExecuted = true;
        var unknownNeverAutoReplayed = true;
        var sameReceiptEveryTime = true;
        var sameReceiptReplayCount = 0;

        using var service = new ExcelMcpService(safetyRoot);
        var sessionId = await ServiceBenchmarkHelpers.CreateSessionAsync(
            service,
            workbookPath,
            context.Configuration.ShowExcel);
        await ServiceBenchmarkHelpers.ConfigureSafetyAsync(
            service,
            sessionId,
            reviewMode: "required",
            checkpointMode: "off");

        var total = context.Configuration.Warmups + context.Configuration.Iterations;
        for (var iteration = 0; iteration < total; iteration++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var values = Enumerable.Range(0, ColumnCount)
                .Select(column => (object?)(iteration * 100_000d + column))
                .ToArray();
            var args = JsonSerializer.Serialize(new
            {
                tableName = "BenchmarkTable",
                rows = new[] { values }
            }, ServiceProtocol.JsonOptions);
            var review = await service.ProcessAsync(new ServiceRequest
            {
                Command = "table.append",
                SessionId = sessionId,
                Args = args,
                ReviewOnly = true,
                Source = "benchmark"
            });
            ServiceBenchmarkHelpers.EnsureSuccess(review, "table.append review");
            var reviewId = BenchmarkContext.GetRequiredString(review.Result, "reviewId");
            var idempotencyKey = $"benchmark-{iteration:D6}";
            var beforeRows = await GetTableRowCountAsync(service, sessionId);

            var executeRequest = new ServiceRequest
            {
                Command = "table.append",
                SessionId = sessionId,
                Args = args,
                ReviewId = reviewId,
                IdempotencyKey = idempotencyKey,
                Source = "benchmark"
            };
            var firstStarted = Stopwatch.GetTimestamp();
            var first = await service.ProcessAsync(executeRequest);
            var firstLatency = BenchmarkContext.ElapsedMilliseconds(firstStarted);
            ServiceBenchmarkHelpers.EnsureSuccess(first, "table.append execute");
            var afterFirstRows = await GetTableRowCountAsync(service, sessionId);

            var retryStarted = Stopwatch.GetTimestamp();
            var retry = await service.ProcessAsync(executeRequest);
            var retryLatency = BenchmarkContext.ElapsedMilliseconds(retryStarted);
            var afterRetryRows = await GetTableRowCountAsync(service, sessionId);
            var duplicateExecutionCount = Math.Max(0, afterRetryRows - afterFirstRows);
            var executesOnce = afterFirstRows == beforeRows + 1 && duplicateExecutionCount == 0;
            executeOnceEveryTime &= executesOnce;
            var sameReceipt = retry.Success && string.Equals(first.Result, retry.Result, StringComparison.Ordinal);
            sameReceiptEveryTime &= sameReceipt;
            if (sameReceipt)
            {
                sameReceiptReplayCount++;
            }

            var changedArgs = JsonSerializer.Serialize(new
            {
                tableName = "BenchmarkTable",
                rows = new[] { Enumerable.Repeat<object?>(-999d, ColumnCount).ToArray() }
            }, ServiceProtocol.JsonOptions);
            var conflictStarted = Stopwatch.GetTimestamp();
            var conflict = await service.ProcessAsync(new ServiceRequest
            {
                Command = "table.append",
                SessionId = sessionId,
                Args = changedArgs,
                ReviewId = reviewId,
                IdempotencyKey = idempotencyKey,
                Source = "benchmark"
            });
            var conflictLatency = BenchmarkContext.ElapsedMilliseconds(conflictStarted);
            var afterConflictRows = await GetTableRowCountAsync(service, sessionId);
            var changedArgumentsBlocked = !conflict.Success && afterConflictRows == afterRetryRows;
            changedArgumentsNeverExecuted &= changedArgumentsBlocked;
            unknownNeverAutoReplayed &= !retry.Success || sameReceipt;

            if (iteration >= context.Configuration.Warmups)
            {
                var receiptBytes = Encoding.UTF8.GetByteCount(first.Result ?? string.Empty);
                observations.Add(new BenchmarkObservation(
                    iteration - context.Configuration.Warmups,
                    "review-id-retry",
                    executesOnce && changedArgumentsBlocked,
                    executesOnce && changedArgumentsBlocked ? null : "A retry or changed-argument request produced an extra table row.",
                    new Dictionary<string, double>
                    {
                        ["first_execution_ms"] = firstLatency,
                        ["duplicate_retry_ms"] = retryLatency,
                        ["duplicate_execution_count"] = duplicateExecutionCount,
                        ["receipt_payload_bytes"] = receiptBytes,
                        ["conflict_detection_ms"] = conflictLatency,
                        ["same_receipt_replayed"] = sameReceipt ? 1 : 0,
                        ["idempotency_key_supported"] = 1
                    },
                    new Dictionary<string, string>
                    {
                        ["retry_outcome"] = retry.Success ? "receipt-replayed" : retry.ErrorCategory ?? "error",
                        ["conflict_outcome"] = conflict.ErrorCategory ?? (conflict.Success ? "unexpected-success" : "error")
                    },
                    retry.Success ? "deduped" : "retry-rejected"));
            }
        }

        await ServiceBenchmarkHelpers.CloseSessionAsync(service, sessionId);
        var plan = BenchmarkPlanCatalog.All.Single(plan => plan.Id == PlanId);
        return ScenarioResult.Create(
            PlanId,
            Name,
            plan.Title,
            observations,
            [
                new BenchmarkInvariant("known_key_executes_once", executeOnceEveryTime, $"Every idempotency key produced exactly one appended row: {executeOnceEveryTime}"),
                new BenchmarkInvariant("same_key_same_receipt", sameReceiptEveryTime, $"Every exact idempotency-key retry returned the original receipt: {sameReceiptEveryTime}"),
                new BenchmarkInvariant("changed_arguments_conflict", changedArgumentsNeverExecuted, $"Changed arguments with the same idempotency key never appended: {changedArgumentsNeverExecuted}"),
                new BenchmarkInvariant("unknown_outcome_not_replayed", unknownNeverAutoReplayed, $"No rejected retry was automatically re-executed: {unknownNeverAutoReplayed}")
            ],
            "Review one table append, execute it with an idempotency key, retry the exact request, then reuse the key with changed arguments and count semantic side effects.",
            [
                $"The idempotency ledger replayed the same receipt in {sameReceiptReplayCount} measured/warmup cases.",
                "Keys are session-scoped and ambiguous timeout/cancellation outcomes are retained as unknown rather than executed again."
            ]);
    }

    private static async Task<int> GetTableRowCountAsync(ExcelMcpService service, string sessionId)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "table.get-data",
            SessionId = sessionId,
            Args = "{\"tableName\":\"BenchmarkTable\",\"visibleOnly\":false}",
            Source = "benchmark"
        });
        ServiceBenchmarkHelpers.EnsureSuccess(response, "table.get-data");
        using var document = JsonDocument.Parse(response.Result!);
        return document.RootElement.GetProperty("rowCount").GetInt32();
    }
}
