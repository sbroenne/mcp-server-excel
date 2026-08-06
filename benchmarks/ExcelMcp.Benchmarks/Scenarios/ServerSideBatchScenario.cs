using System.Diagnostics;
using System.Text.Json;
using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.Benchmarks.Scenarios;

internal sealed class ServerSideBatchScenario : IBenchmarkScenario
{
    private const int OperationCount = 20;

    public string PlanId => "04";

    public string Name => "server-side-batch";

    public async Task<ScenarioResult> RunAsync(BenchmarkContext context, CancellationToken cancellationToken)
    {
        var masterPath = context.CreateWorkingPath("batch-master");
        BenchmarkContext.CreateDataWorkbook(masterPath, rows: OperationCount + 2, columns: 1);
        var workbookPath = context.CopyWorkbook(masterPath, "batch-workload");
        var safetyRoot = context.CreateSafetyRoot("batch");
        var observations = new List<BenchmarkObservation>();
        var exactValuesEveryTime = true;
        var failureIndexReported = true;
        var sessionCleaned = false;
        var serverBatchSupported = false;
        int? excelProcessId = null;

        using (var service = new ExcelMcpService(safetyRoot))
        {
            var sessionId = await ServiceBenchmarkHelpers.CreateSessionAsync(
                service,
                workbookPath,
                context.Configuration.ShowExcel);
            await ServiceBenchmarkHelpers.ConfigureSafetyAsync(
                service,
                sessionId,
                reviewMode: "off",
                checkpointMode: "off",
                journalMode: "off",
                verificationMode: "off");
            excelProcessId = service.SessionManager.GetSession(sessionId)?.ExcelProcessId;
            serverBatchSupported = await DetectServerBatchAsync(service, sessionId, cancellationToken);

            var total = context.Configuration.Warmups + context.Configuration.Iterations;
            for (var index = 0; index < total; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var operations = CreateWriteOperations(index);
                var started = Stopwatch.GetTimestamp();
                var execution = serverBatchSupported
                    ? await ExecuteServerBatchAsync(service, sessionId, operations)
                    : await ExecuteSeparateAsync(service, sessionId, operations, stopOnError: true);
                var elapsed = BenchmarkContext.ElapsedMilliseconds(started);
                var values = await ReadValuesAsync(service, sessionId, OperationCount);
                var expected = Enumerable.Range(0, OperationCount).Select(item => index * 10_000d + item).ToArray();
                var exactValues = values.SequenceEqual(expected);
                exactValuesEveryTime &= exactValues;

                if (index >= context.Configuration.Warmups)
                {
                    observations.Add(new BenchmarkObservation(
                        index - context.Configuration.Warmups,
                        "twenty-ordered-writes",
                        execution.Success && exactValues,
                        execution.Error,
                        new Dictionary<string, double>
                        {
                            ["batch_latency_ms"] = elapsed,
                            ["operations_per_second"] = OperationCount / (elapsed / 1000d),
                            ["request_count"] = execution.RequestCount,
                            ["payload_bytes"] = execution.PayloadBytes,
                            ["token_estimate"] = BenchmarkContext.EstimateTokensFromUtf8Bytes(execution.PayloadBytes),
                            ["candidate_batch_supported"] = serverBatchSupported ? 1 : 0
                        },
                        new Dictionary<string, string>
                        {
                            ["implementation"] = serverBatchSupported ? "session.batch" : "separate-requests",
                            ["operation_count"] = OperationCount.ToString(System.Globalization.CultureInfo.InvariantCulture)
                        },
                        execution.Success ? "completed" : "failed"));
                }
            }

            var failureOperations = new List<ServiceOperation>
            {
                new("range.set-values", new { sheetName = "Data", rangeAddress = "A2", values = new object?[][] { [1] } }),
                new("range.set-values", new { sheetName = "MissingSheet", rangeAddress = "A1", values = new object?[][] { [2] } }),
                new("range.set-values", new { sheetName = "Data", rangeAddress = "A3", values = new object?[][] { [3] } })
            };
            var failureResult = serverBatchSupported
                ? await ExecuteServerBatchAsync(service, sessionId, failureOperations)
                : await ExecuteSeparateAsync(service, sessionId, failureOperations, stopOnError: true);
            failureIndexReported = failureResult.FailedIndex == 1;
            observations.Add(new BenchmarkObservation(
                context.Configuration.Iterations,
                "failure-index",
                failureIndexReported,
                failureIndexReported ? null : $"Expected failed operation index 1, got {failureResult.FailedIndex?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "none"}.",
                new Dictionary<string, double>
                {
                    ["failure_index_reported"] = failureIndexReported ? 1 : 0,
                    ["request_count"] = failureResult.RequestCount
                },
                Outcome: failureResult.Success ? "unexpected-success" : "expected-failure"));

            await ServiceBenchmarkHelpers.CloseSessionAsync(service, sessionId);
        }

        if (excelProcessId.HasValue)
        {
            sessionCleaned = BenchmarkContext.WaitForProcessExit(
                excelProcessId.Value,
                TimeSpan.FromSeconds(15),
                out _);
        }

        var protocol = await ProtocolFootprintProbe.RunAsync(context.Configuration.Iterations, workbookPath, cancellationToken);
        observations.AddRange(protocol.Observations);

        var plan = BenchmarkPlanCatalog.All.Single(plan => plan.Id == PlanId);
        return ScenarioResult.Create(
            PlanId,
            Name,
            plan.Title,
            observations,
            [
                new BenchmarkInvariant("operation_order", exactValuesEveryTime, $"Every final range matched request order: {exactValuesEveryTime}"),
                new BenchmarkInvariant("no_lost_operation", exactValuesEveryTime, $"All {OperationCount} target cells were present after every workload: {exactValuesEveryTime}"),
                new BenchmarkInvariant("no_duplicate_operation", exactValuesEveryTime, $"Every target cell held exactly its expected unique value: {exactValuesEveryTime}"),
                new BenchmarkInvariant("failure_index_reported", failureIndexReported, $"The synthetic failure was attributed to index 1: {failureIndexReported}"),
                new BenchmarkInvariant("session_cleanup", sessionCleaned, $"Owned Excel process exited after session.close: {sessionCleaned}"),
                new BenchmarkInvariant("mcp_tool_surface", protocol.ToolCount > 0 && protocol.ToolCallSucceeded, $"MCP listed {protocol.ToolCount} tools and the fixed tool call returned success=true: {protocol.ToolCallSucceeded}")
            ],
            $"{OperationCount} ordered one-cell writes through the service seam, plus actual MCP initialize/tools-list/tools-call wire capture.",
            [
                serverBatchSupported
                    ? "The candidate session.batch contract was detected and used."
                    : "Current baseline has no session.batch contract, so the identical workload used separate service requests.",
                "Token counts are deterministic ceil(UTF-8 wire bytes / 4) estimates. Raw bytes and schema hashes are preserved for exact comparison."
            ]);
    }

    private static ServiceOperation[] CreateWriteOperations(int iteration) =>
        Enumerable.Range(0, OperationCount)
            .Select(index => new ServiceOperation(
                "range.set-values",
                new
                {
                    sheetName = "Data",
                    rangeAddress = $"A{index + 2}",
                    values = new object?[][] { [iteration * 10_000d + index] }
                }))
            .ToArray();

    private static async Task<bool> DetectServerBatchAsync(
        ExcelMcpService service,
        string sessionId,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var probe = await ExecuteServerBatchAsync(
            service,
            sessionId,
            [new ServiceOperation("range.get-values", new { sheetName = "Data", rangeAddress = "A1" })]);
        if (probe.Success)
        {
            return true;
        }

        if (string.Equals(probe.Error, "Unknown session action: batch", StringComparison.Ordinal))
        {
            return false;
        }

        throw new InvalidOperationException($"session.batch capability probe failed ambiguously: {probe.Error ?? "no error details"}");
    }

    private static async Task<BatchExecution> ExecuteSeparateAsync(
        ExcelMcpService service,
        string sessionId,
        IReadOnlyList<ServiceOperation> operations,
        bool stopOnError)
    {
        long payloadBytes = 0;
        for (var index = 0; index < operations.Count; index++)
        {
            var operation = operations[index];
            var request = new ServiceRequest
            {
                Command = operation.Command,
                SessionId = sessionId,
                Args = JsonSerializer.Serialize(operation.Args, ServiceProtocol.JsonOptions),
                Source = "benchmark"
            };
            var response = await service.ProcessAsync(request);
            payloadBytes += ServiceBenchmarkHelpers.SerializedPayloadBytes(request, response);
            if (!response.Success && stopOnError)
            {
                return new BatchExecution(false, index + 1, payloadBytes, index, response.ErrorMessage);
            }
        }

        return new BatchExecution(true, operations.Count, payloadBytes, null, null);
    }

    private static async Task<BatchExecution> ExecuteServerBatchAsync(
        ExcelMcpService service,
        string sessionId,
        IReadOnlyList<ServiceOperation> operations)
    {
        var request = new ServiceRequest
        {
            Command = "session.batch",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                operations = operations.Select(operation => new { command = operation.Command, args = operation.Args }),
                stopOnError = true
            }, ServiceProtocol.JsonOptions),
            Source = "benchmark"
        };
        var response = await service.ProcessAsync(request);
        int? failedIndex = null;
        if (!response.Success && !string.IsNullOrWhiteSpace(response.Result))
        {
            using var document = JsonDocument.Parse(response.Result);
            if (document.RootElement.TryGetProperty("failedIndex", out var property) && property.TryGetInt32(out var value))
            {
                failedIndex = value;
            }
        }

        return new BatchExecution(
            response.Success,
            1,
            ServiceBenchmarkHelpers.SerializedPayloadBytes(request, response),
            failedIndex,
            response.ErrorMessage);
    }

    private static async Task<double[]> ReadValuesAsync(ExcelMcpService service, string sessionId, int count)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.get-values",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new { sheetName = "Data", rangeAddress = $"A2:A{count + 1}" }, ServiceProtocol.JsonOptions),
            Source = "benchmark"
        });
        ServiceBenchmarkHelpers.EnsureSuccess(response, "range.get-values verification");
        using var document = JsonDocument.Parse(response.Result!);
        return document.RootElement.GetProperty("values")
            .EnumerateArray()
            .Select(row => row[0].GetDouble())
            .ToArray();
    }

    private sealed record ServiceOperation(string Command, object Args);

    private sealed record BatchExecution(bool Success, int RequestCount, long PayloadBytes, int? FailedIndex, string? Error);
}
