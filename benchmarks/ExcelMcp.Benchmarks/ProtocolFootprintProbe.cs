using System.Diagnostics;
using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;

namespace Sbroenne.ExcelMcp.Benchmarks;

internal sealed record ProtocolFootprintResult(
    IReadOnlyList<BenchmarkObservation> Observations,
    int ToolCount,
    bool ToolCallSucceeded,
    long InitializationPayloadBytes);

internal static class ProtocolFootprintProbe
{
    private static readonly TimeSpan TestTransportShutdownTimeout = TimeSpan.FromSeconds(30);
    private static readonly HashSet<string> UnknownOutcomeCategories = new(StringComparer.OrdinalIgnoreCase)
    {
        "Timeout",
        "Cancelled",
        "Canceled",
        "ExcelProcessDied",
        "IdempotencyUnknownOutcome",
        "IdempotencyInProgress",
        "JournalPersistenceFailed",
        "AbortedUnknown",
        "UnknownOutcome",
        "SessionInterrupted",
        "ServerShutdown"
    };

    /// <summary>
    /// Runs one complete, public MCP workbook-edit workflow on a fresh transport.
    /// The fixture workbook must already exist; setup is intentionally outside the measured prompt.
    /// </summary>
    internal static async Task<PromptWorkflowRunResult> RunPromptWorkflowAsync(
        PromptWorkflowVariant variant,
        string workbookPath,
        IReadOnlyList<PromptWorkflowWrite> writes,
        bool showExcel,
        CancellationToken cancellationToken)
    {
        var clientToServer = new Pipe();
        var serverToClient = new Pipe();
        await using var clientInput = new CaptureStream(clientToServer.Writer.AsStream(), captureWrites: true, leaveOpen: true);
        await using var clientOutput = new CaptureStream(serverToClient.Reader.AsStream(), captureWrites: false, leaveOpen: true);
        McpClient? client = null;
        Task<int>? serverTask = null;
        string? sessionId = null;
        string description = "{}";
        IReadOnlyList<double> values = [];
        string? error = null;
        var knownOutcome = true;
        var sessionClosed = false;
        var toolCallCount = 0;
        long toolCallRequestBytes = 0;
        long toolCallResponseBytes = 0;
        long initializeRequestBytes = 0;
        long initializeResponseBytes = 0;
        long toolsListRequestBytes = 0;
        long toolsListResponseBytes = 0;
        var promptToCompletionMilliseconds = 0d;
        var openDescribeMilliseconds = 0d;
        var executionMilliseconds = 0d;
        var verificationMilliseconds = 0d;
        var excludedCorrectnessAuditMilliseconds = 0d;
        var promptStarted = Stopwatch.GetTimestamp();
        Exception? teardownFailure = null;
        var serverStopped = false;

        try
        {
            Sbroenne.ExcelMcp.McpServer.Program.ConfigureTestTransport(clientToServer, serverToClient);
            serverTask = Sbroenne.ExcelMcp.McpServer.Program.Main(
                ["--tool-profile", "copilot-compact"]);

            var initializeRequestStart = clientInput.CapturedLength;
            var initializeResponseStart = clientOutput.CapturedLength;
            client = await McpClient.CreateAsync(
                new StreamClientTransport(serverInput: clientInput, serverOutput: clientOutput),
                clientOptions: new McpClientOptions
                {
                    ClientInfo = new() { Name = "ExcelMcpPromptBenchmark", Version = "1.0.0" },
                    InitializationTimeout = TimeSpan.FromSeconds(30)
                },
                cancellationToken: cancellationToken);
            initializeRequestBytes = clientInput.CapturedLength - initializeRequestStart;
            initializeResponseBytes = clientOutput.CapturedLength - initializeResponseStart;

            var toolsListRequestStart = clientInput.CapturedLength;
            var toolsListResponseStart = clientOutput.CapturedLength;
            var tools = await client.ListToolsAsync(cancellationToken: cancellationToken);
            toolsListRequestBytes = clientInput.CapturedLength - toolsListRequestStart;
            toolsListResponseBytes = clientOutput.CapturedLength - toolsListResponseStart;
            EnsureWorkflowSurface(tools, variant);

            async Task<string> CallToolJsonAsync(string toolName, Dictionary<string, object?> arguments)
            {
                var requestStart = clientInput.CapturedLength;
                var responseStart = clientOutput.CapturedLength;
                var result = await client.CallToolAsync(toolName, arguments, cancellationToken: cancellationToken);
                toolCallRequestBytes += clientInput.CapturedLength - requestStart;
                toolCallResponseBytes += clientOutput.CapturedLength - responseStart;
                toolCallCount++;
                return GetToolText(result);
            }

            async Task<string> CallToolJsonForCorrectnessAuditAsync(
                string toolName,
                Dictionary<string, object?> arguments)
            {
                // This independent read proves workbook correctness, but it is deliberately
                // excluded from public-prompt call, byte, token, and latency metrics. The
                // optimized client workflow ends with the bounded verification receipt.
                var result = await client.CallToolAsync(toolName, arguments, cancellationToken: cancellationToken);
                return GetToolText(result);
            }

            var openDescribeStarted = Stopwatch.GetTimestamp();
            if (variant == PromptWorkflowVariant.ExecutePlanAndOpenDescribe)
            {
                var openDescribe = await CallToolJsonAsync("workflow", new Dictionary<string, object?>
                {
                    ["action"] = "open-and-describe",
                    ["file_path"] = workbookPath,
                    ["show"] = showExcel,
                    ["preview_rows"] = 3,
                    ["preview_columns"] = 3
                });
                EnsureSuccess(openDescribe, "workflow.open-and-describe");
                sessionId = GetRequiredSessionId(openDescribe);
                description = CompactJson(openDescribe);
            }
            else
            {
                var open = await CallToolJsonAsync("file", new Dictionary<string, object?>
                {
                    ["action"] = "open",
                    ["path"] = workbookPath,
                    ["show"] = showExcel
                });
                EnsureSuccess(open, "file.open");
                sessionId = GetRequiredSessionId(open);

                var sheets = await CallToolJsonAsync("worksheet", new Dictionary<string, object?>
                {
                    ["action"] = "list",
                    ["session_id"] = sessionId
                });
                EnsureSuccess(sheets, "worksheet.list");

                var usedRange = await CallToolJsonAsync("range", new Dictionary<string, object?>
                {
                    ["action"] = "get-used-range",
                    ["session_id"] = sessionId,
                    ["sheet_name"] = "Data"
                });
                EnsureSuccess(usedRange, "range.get-used-range");
                description = BuildLegacyDescription(sheets, usedRange);
            }
            openDescribeMilliseconds = BenchmarkContext.ElapsedMilliseconds(openDescribeStarted);

            var executionStarted = Stopwatch.GetTimestamp();
            if (variant == PromptWorkflowVariant.Legacy)
            {
                foreach (var write in writes)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var writeResult = await CallToolJsonAsync("range", new Dictionary<string, object?>
                    {
                        ["action"] = "set-values",
                        ["session_id"] = sessionId,
                        ["sheet_name"] = "Data",
                        ["range_address"] = write.Address,
                        ["values"] = new object?[][] { [write.Value] }
                    });
                    EnsureSuccess(writeResult, "range.set-values");
                }
            }
            else
            {
                var executePlan = await CallToolJsonAsync("workflow", new Dictionary<string, object?>
                {
                    ["action"] = "execute-plan",
                    ["session_id"] = sessionId,
                    ["operations"] = writes.Select(write => new Dictionary<string, object?>
                    {
                        ["command"] = "range.set-values",
                        ["args"] = new Dictionary<string, object?>
                        {
                            ["sheetName"] = "Data",
                            ["rangeAddress"] = write.Address,
                            ["values"] = new object?[][] { [write.Value] }
                        }
                    }).ToArray(),
                    ["stop_on_error"] = true,
                    ["verify_sheet_name"] = "Data",
                    ["verify_range_address"] = $"A2:A{writes.Count + 1}"
                });
                EnsureSuccess(executePlan, "workflow.execute-plan");
                EnsureFinalVerificationReceipt(executePlan, writes);
            }
            executionMilliseconds = BenchmarkContext.ElapsedMilliseconds(executionStarted);

            if (variant == PromptWorkflowVariant.Legacy)
            {
                var verificationStarted = Stopwatch.GetTimestamp();
                var verification = await CallToolJsonAsync("range", new Dictionary<string, object?>
                {
                    ["action"] = "get-values",
                    ["session_id"] = sessionId,
                    ["sheet_name"] = "Data",
                    ["range_address"] = $"A2:A{writes.Count + 1}"
                });
                EnsureSuccess(verification, "range.get-values verification");
                values = ParseValues(verification);
                verificationMilliseconds = BenchmarkContext.ElapsedMilliseconds(verificationStarted);
            }
            else
            {
                var auditStarted = Stopwatch.GetTimestamp();
                try
                {
                    var audit = await CallToolJsonForCorrectnessAuditAsync("range", new Dictionary<string, object?>
                    {
                        ["action"] = "get-values",
                        ["session_id"] = sessionId,
                        ["sheet_name"] = "Data",
                        ["range_address"] = $"A2:A{writes.Count + 1}"
                    });
                    EnsureSuccess(audit, "unmeasured range.get-values correctness audit");
                    values = ParseValues(audit);
                }
                finally
                {
                    excludedCorrectnessAuditMilliseconds += BenchmarkContext.ElapsedMilliseconds(auditStarted);
                }
            }
        }
        catch (OperationCanceledException exception) when (!cancellationToken.IsCancellationRequested)
        {
            error = $"{exception.GetType().Name}: {exception.Message}";
            knownOutcome = !IsUnknownOutcome(error);
        }
        catch (Exception exception) when (exception is not OperationCanceledException)
        {
            error = $"{exception.GetType().Name}: {exception.Message}";
            knownOutcome = !IsUnknownOutcome(error);
        }
        finally
        {
            void RecordTeardownFailure(Exception exception) => teardownFailure ??= exception;

            if (client is not null && !string.IsNullOrWhiteSpace(sessionId))
            {
                try
                {
                    var requestStart = clientInput.CapturedLength;
                    var responseStart = clientOutput.CapturedLength;
                    var close = await client.CallToolAsync("file", new Dictionary<string, object?>
                    {
                        ["action"] = "close",
                        ["session_id"] = sessionId,
                        ["save"] = false
                    }, cancellationToken: CancellationToken.None);
                    toolCallRequestBytes += clientInput.CapturedLength - requestStart;
                    toolCallResponseBytes += clientOutput.CapturedLength - responseStart;
                    toolCallCount++;
                    var closeJson = GetToolText(close);
                    if (IsSuccessfulToolResult(closeJson))
                    {
                        sessionClosed = true;
                    }
                    else
                    {
                        error ??= $"file.close failed: {closeJson}";
                        knownOutcome &= !IsUnknownOutcome(closeJson);
                    }
                }
                catch (Exception exception)
                {
                    error ??= $"Cleanup failed: {exception.GetType().Name}: {exception.Message}";
                    knownOutcome &= !IsUnknownOutcome(exception.Message);
                }
            }

            // The prompt ends after its final public tool call. Transport teardown is hygiene,
            // not user-visible prompt work, so it is deliberately excluded from this latency.
            promptToCompletionMilliseconds = Math.Max(
                0d,
                BenchmarkContext.ElapsedMilliseconds(promptStarted) - excludedCorrectnessAuditMilliseconds);

            try
            {
                if (client is not null)
                {
                    await client.DisposeAsync();
                }
            }
            catch (Exception exception)
            {
                RecordTeardownFailure(new InvalidOperationException("MCP client disposal failed during benchmark teardown.", exception));
            }

            try
            {
                Sbroenne.ExcelMcp.McpServer.Program.RequestTestTransportShutdown();
            }
            catch (Exception exception)
            {
                RecordTeardownFailure(new InvalidOperationException("MCP server shutdown request failed during benchmark teardown.", exception));
            }

            try
            {
                await TryCompleteAsync(clientToServer.Writer);
                await TryCompleteAsync(serverToClient.Reader);
            }
            catch (Exception exception)
            {
                RecordTeardownFailure(new InvalidOperationException("MCP transport completion failed during benchmark teardown.", exception));
            }

            if (serverTask is null)
            {
                serverStopped = true;
            }
            else
            {
                try
                {
                    await serverTask.WaitAsync(TestTransportShutdownTimeout, CancellationToken.None);
                    serverStopped = true;
                }
                catch (TimeoutException exception)
                {
                    RecordTeardownFailure(new TimeoutException(
                        "MCP benchmark server did not stop after shutdown and pipe completion; refusing to run another case on a contaminated transport.",
                        exception));
                }
                catch (Exception exception)
                {
                    serverStopped = serverTask.IsCompleted;
                    RecordTeardownFailure(new InvalidOperationException("MCP benchmark server faulted during teardown.", exception));
                }
            }

            try
            {
                await TryCompleteAsync(clientToServer.Reader);
                await TryCompleteAsync(serverToClient.Writer);
            }
            catch (Exception exception)
            {
                RecordTeardownFailure(new InvalidOperationException("MCP transport final completion failed during benchmark teardown.", exception));
            }
            finally
            {
                if (serverStopped)
                {
                    try
                    {
                        Sbroenne.ExcelMcp.McpServer.Program.ResetTestTransport();
                    }
                    catch (Exception exception)
                    {
                        RecordTeardownFailure(new InvalidOperationException("MCP test transport reset failed during benchmark teardown.", exception));
                    }
                }
                else
                {
                    RecordTeardownFailure(new InvalidOperationException(
                        "MCP test transport remains configured because its server task is still live after bounded shutdown; the benchmark fails closed to prevent another case from reusing contaminated static state."));
                }
            }
        }

        if (teardownFailure is not null)
        {
            throw new InvalidOperationException("MCP benchmark transport teardown failed.", teardownFailure);
        }

        return new PromptWorkflowRunResult(
            variant,
            error is null,
            error,
            knownOutcome,
            sessionClosed,
            description,
            values,
            promptToCompletionMilliseconds,
            openDescribeMilliseconds,
            executionMilliseconds,
            verificationMilliseconds,
            toolCallCount,
            new McpWireByteBreakdown(
                initializeRequestBytes,
                initializeResponseBytes,
                toolsListRequestBytes,
                toolsListResponseBytes,
                toolCallRequestBytes,
                toolCallResponseBytes));
    }

    public static async Task<ProtocolFootprintResult> RunAsync(
        int iterations,
        string testWorkbookPath,
        CancellationToken cancellationToken)
    {
        var clientToServer = new Pipe();
        var serverToClient = new Pipe();
        await using var clientInput = new CaptureStream(clientToServer.Writer.AsStream(), captureWrites: true, leaveOpen: true);
        await using var clientOutput = new CaptureStream(serverToClient.Reader.AsStream(), captureWrites: false, leaveOpen: true);
        McpClient? client = null;
        Task<int>? serverTask = null;

        try
        {
            Sbroenne.ExcelMcp.McpServer.Program.ConfigureTestTransport(clientToServer, serverToClient);
            serverTask = Sbroenne.ExcelMcp.McpServer.Program.Main([]);
            client = await McpClient.CreateAsync(
                new StreamClientTransport(serverInput: clientInput, serverOutput: clientOutput),
                clientOptions: new McpClientOptions
                {
                    ClientInfo = new() { Name = "ExcelMcpBenchmark", Version = "1.0.0" },
                    InitializationTimeout = TimeSpan.FromSeconds(30)
                },
                cancellationToken: cancellationToken);

            var initializationBytes = clientInput.CapturedLength + clientOutput.CapturedLength;
            var observations = new List<BenchmarkObservation>(iterations);
            var toolCount = 0;
            var toolCallSucceeded = true;

            for (var iteration = 0; iteration < iterations; iteration++)
            {
                var requestStart = clientInput.CapturedLength;
                var responseStart = clientOutput.CapturedLength;
                var tools = await client.ListToolsAsync(cancellationToken: cancellationToken);
                toolCount = tools.Count;
                var requestBytes = clientInput.CapturedLength - requestStart;
                var responseBytes = clientOutput.CapturedLength - responseStart;
                var schemaBytes = requestBytes + responseBytes;
                var schemaHash = BenchmarkContext.Sha256(Convert.ToBase64String(clientOutput.GetCapturedBytes(responseStart)));

                requestStart = clientInput.CapturedLength;
                responseStart = clientOutput.CapturedLength;
                var callResult = await client.CallToolAsync(
                    "file",
                    new Dictionary<string, object?>
                    {
                        ["action"] = "test",
                        ["path"] = testWorkbookPath
                    },
                    cancellationToken: cancellationToken);
                var callBytes = (clientInput.CapturedLength - requestStart) + (clientOutput.CapturedLength - responseStart);
                var callSucceeded = IsSuccessfulJsonResponse(callResult);
                toolCallSucceeded &= callSucceeded;

                observations.Add(new BenchmarkObservation(
                    iteration,
                    "mcp-wire-footprint",
                    tools.Count > 0 && callSucceeded,
                    null,
                    new Dictionary<string, double>
                    {
                        ["mcp_schema_payload_bytes"] = schemaBytes,
                        ["mcp_schema_token_estimate"] = BenchmarkContext.EstimateTokensFromUtf8Bytes(schemaBytes),
                        ["schema_response_bytes"] = responseBytes,
                        ["schema_request_bytes"] = requestBytes,
                        ["tool_call_payload_bytes"] = callBytes,
                        ["tool_call_token_estimate"] = BenchmarkContext.EstimateTokensFromUtf8Bytes(callBytes),
                        ["tool_count"] = tools.Count,
                        ["initialization_payload_bytes"] = initializationBytes
                    },
                    new Dictionary<string, string>
                    {
                        ["token_measurement"] = "ceil(utf8-wire-bytes/4); deterministic estimate, not model-specific",
                        ["schema_response_sha256"] = schemaHash
                    },
                    "protocol-success"));
            }

            return new ProtocolFootprintResult(observations, toolCount, toolCallSucceeded, initializationBytes);
        }
        finally
        {
            if (client is not null)
            {
                await client.DisposeAsync();
            }

            Sbroenne.ExcelMcp.McpServer.Program.RequestTestTransportShutdown();
            await TryCompleteAsync(clientToServer.Writer);
            await TryCompleteAsync(serverToClient.Reader);

            if (serverTask is not null)
            {
                try
                {
                    await serverTask.WaitAsync(TimeSpan.FromSeconds(30), cancellationToken);
                }
                catch (OperationCanceledException)
                {
                }
                catch (TimeoutException)
                {
                }
            }

            await TryCompleteAsync(clientToServer.Reader);
            await TryCompleteAsync(serverToClient.Writer);
            Sbroenne.ExcelMcp.McpServer.Program.ResetTestTransport();
        }
    }

    private static async Task TryCompleteAsync(PipeReader reader)
    {
        try
        {
            await reader.CompleteAsync();
        }
        catch (InvalidOperationException)
        {
        }
    }

    private static bool IsSuccessfulJsonResponse(CallToolResult result)
    {
        var text = result.Content.OfType<TextContentBlock>().FirstOrDefault()?.Text;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        try
        {
            using var document = JsonDocument.Parse(text);
            return document.RootElement.TryGetProperty("success", out var success) && success.GetBoolean();
        }
        catch (JsonException)
        {
            return false;
        }
    }

    private static string GetToolText(CallToolResult result) =>
        result.Content.OfType<TextContentBlock>().FirstOrDefault()?.Text
            ?? throw new InvalidDataException("MCP tool call returned no text result.");

    private static void EnsureSuccess(string result, string operation)
    {
        if (!IsSuccessfulToolResult(result))
        {
            throw new InvalidDataException($"{operation} failed: {result}");
        }
    }

    internal static bool IsSuccessfulToolResult(string result)
    {
        try
        {
            using var document = JsonDocument.Parse(result);
            if (document.RootElement.TryGetProperty("success", out var success) &&
                success.ValueKind is JsonValueKind.True or JsonValueKind.False)
            {
                return success.GetBoolean();
            }

            return document.RootElement.TryGetProperty("outcome", out var outcome) &&
                string.Equals(outcome.GetString(), "completed", StringComparison.OrdinalIgnoreCase);
        }
        catch (JsonException)
        {
            return false;
        }
    }

    private static string GetRequiredSessionId(string result)
    {
        using var document = JsonDocument.Parse(result);
        var root = document.RootElement;
        if (root.TryGetProperty("session_id", out var snakeCase) && !string.IsNullOrWhiteSpace(snakeCase.GetString()))
        {
            return snakeCase.GetString()!;
        }

        if (root.TryGetProperty("sessionId", out var camelCase) && !string.IsNullOrWhiteSpace(camelCase.GetString()))
        {
            return camelCase.GetString()!;
        }

        throw new InvalidDataException("MCP response did not include a session identifier.");
    }

    private static string BuildLegacyDescription(string sheets, string usedRange) =>
        JsonSerializer.Serialize(new
        {
            sheets = JsonSerializer.Deserialize<JsonElement>(sheets),
            usedRange = JsonSerializer.Deserialize<JsonElement>(usedRange)
        }, Sbroenne.ExcelMcp.ComInterop.ServiceClient.ServiceProtocol.JsonOptions);

    private static string CompactJson(string json)
    {
        using var document = JsonDocument.Parse(json);
        return JsonSerializer.Serialize(document.RootElement, Sbroenne.ExcelMcp.ComInterop.ServiceClient.ServiceProtocol.JsonOptions);
    }

    private static double[] ParseValues(string result)
    {
        using var document = JsonDocument.Parse(result);
        return document.RootElement.GetProperty("values")
            .EnumerateArray()
            .Select(row => row[0].GetDouble())
            .ToArray();
    }

    internal static void EnsureFinalVerificationReceipt(
        string result,
        IReadOnlyList<PromptWorkflowWrite> writes)
    {
        using var document = JsonDocument.Parse(result);
        var root = document.RootElement;
        if (!root.TryGetProperty("verification", out var verification) ||
            !string.Equals(verification.GetProperty("status").GetString(), "verified", StringComparison.Ordinal) ||
            !string.Equals(verification.GetProperty("sheetName").GetString(), "Data", StringComparison.Ordinal) ||
            verification.GetProperty("rowCount").GetInt32() != writes.Count ||
            verification.GetProperty("columnCount").GetInt32() != 1 ||
            verification.GetProperty("cellCount").GetInt64() != writes.Count ||
            verification.GetProperty("inspectedCellCount").GetInt32() != writes.Count ||
            verification.GetProperty("nonEmptyCellCount").GetInt32() != writes.Count ||
            verification.GetProperty("formulaCellCount").GetInt32() != 0)
        {
            throw new InvalidDataException("workflow.execute-plan did not return the expected complete final-range verification receipt.");
        }

        string? fingerprint = verification.GetProperty("fingerprint").GetString();
        if (fingerprint?.Length != 64 || !fingerprint.All(Uri.IsHexDigit))
        {
            throw new InvalidDataException("workflow.execute-plan returned an invalid verification fingerprint.");
        }

        var preview = verification.GetProperty("preview");
        int previewCount = Math.Min(2, writes.Count);
        if (preview.GetArrayLength() != previewCount)
        {
            throw new InvalidDataException("workflow.execute-plan returned an unexpected bounded verification preview.");
        }

        for (int index = 0; index < previewCount; index++)
        {
            if (preview[index].GetArrayLength() != 1 ||
                preview[index][0].GetDouble() != writes[index].Value)
            {
                throw new InvalidDataException("workflow.execute-plan verification preview did not match the requested writes.");
            }
        }
    }

    private static bool IsUnknownOutcome(string? message) =>
        IsConservativeUnknownOutcome(null, message);

    internal static bool IsConservativeUnknownOutcome(string? errorCategory, string? message)
    {
        if (IsKnownNotExecuted(errorCategory, message))
        {
            return false;
        }

        var category = errorCategory?.Trim();
        if (category is not null && UnknownOutcomeCategories.Contains(category))
        {
            return true;
        }

        return message?.Contains("timeout", StringComparison.OrdinalIgnoreCase) == true ||
            message?.Contains("timed out", StringComparison.OrdinalIgnoreCase) == true ||
            message?.Contains("cancelled", StringComparison.OrdinalIgnoreCase) == true ||
            message?.Contains("canceled", StringComparison.OrdinalIgnoreCase) == true ||
            message?.Contains("excelprocessdied", StringComparison.OrdinalIgnoreCase) == true ||
            message?.Contains("excel process died", StringComparison.OrdinalIgnoreCase) == true ||
            message?.Contains("no longer running", StringComparison.OrdinalIgnoreCase) == true ||
            message?.Contains("disconnected", StringComparison.OrdinalIgnoreCase) == true ||
            message?.Contains("idempotencyunknownoutcome", StringComparison.OrdinalIgnoreCase) == true ||
            message?.Contains("idempotency unknown outcome", StringComparison.OrdinalIgnoreCase) == true ||
            message?.Contains("idempotency in progress", StringComparison.OrdinalIgnoreCase) == true ||
            message?.Contains("journalpersistencefailed", StringComparison.OrdinalIgnoreCase) == true ||
            message?.Contains("journal persistence failed", StringComparison.OrdinalIgnoreCase) == true ||
            message?.Contains("abortedunknown", StringComparison.OrdinalIgnoreCase) == true ||
            message?.Contains("aborted unknown", StringComparison.OrdinalIgnoreCase) == true ||
            message?.Contains("server shutdown", StringComparison.OrdinalIgnoreCase) == true ||
            message?.Contains("unknown outcome", StringComparison.OrdinalIgnoreCase) == true ||
            message?.Contains("outcome is unknown", StringComparison.OrdinalIgnoreCase) == true ||
            message?.Contains("outcome unknown", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static bool IsKnownNotExecuted(string? errorCategory, string? message) =>
        string.Equals(errorCategory, "TimeoutBeforeExecution", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(errorCategory, "CancelledBeforeExecution", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(errorCategory, "CheckpointFailed", StringComparison.OrdinalIgnoreCase) ||
        message?.Contains("before execution", StringComparison.OrdinalIgnoreCase) == true;

    private static void EnsureWorkflowSurface(IList<McpClientTool> tools, PromptWorkflowVariant variant)
    {
        if (variant != PromptWorkflowVariant.Legacy && !tools.Any(tool => tool.Name == "workflow"))
        {
            throw new InvalidOperationException("The workflow MCP tool is unavailable for this benchmark variant.");
        }
    }

    private static async Task TryCompleteAsync(PipeWriter writer)
    {
        try
        {
            await writer.CompleteAsync();
        }
        catch (InvalidOperationException)
        {
        }
    }
}

internal enum PromptWorkflowVariant
{
    Legacy,
    ExecutePlanOnly,
    ExecutePlanAndOpenDescribe
}

internal sealed record PromptWorkflowWrite(string Address, double Value);

internal sealed record McpWireByteBreakdown(
    long InitializeRequestBytes,
    long InitializeResponseBytes,
    long ToolsListRequestBytes,
    long ToolsListResponseBytes,
    long ToolCallRequestBytes,
    long ToolCallResponseBytes)
{
    public long TotalBytes => InitializeRequestBytes + InitializeResponseBytes +
        ToolsListRequestBytes + ToolsListResponseBytes + ToolCallRequestBytes + ToolCallResponseBytes;
}

internal sealed record PromptWorkflowRunResult(
    PromptWorkflowVariant Variant,
    bool Success,
    string? Error,
    bool KnownOutcome,
    bool SessionClosed,
    string Description,
    IReadOnlyList<double> Values,
    double PromptToCompletionMilliseconds,
    double OpenDescribeMilliseconds,
    double ExecutionMilliseconds,
    double VerificationMilliseconds,
    int ToolCallCount,
    McpWireByteBreakdown WireBytes);

internal sealed class CaptureStream : Stream
{
    private readonly Stream _inner;
    private readonly bool _captureWrites;
    private readonly bool _leaveOpen;
    private readonly MemoryStream _capture = new();
    private readonly object _sync = new();

    public CaptureStream(Stream inner, bool captureWrites, bool leaveOpen)
    {
        _inner = inner;
        _captureWrites = captureWrites;
        _leaveOpen = leaveOpen;
    }

    public long CapturedLength
    {
        get
        {
            lock (_sync)
            {
                return _capture.Length;
            }
        }
    }

    public override bool CanRead => _inner.CanRead;

    public override bool CanSeek => false;

    public override bool CanWrite => _inner.CanWrite;

    public override long Length => throw new NotSupportedException();

    public override long Position
    {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    public byte[] GetCapturedBytes(long start)
    {
        lock (_sync)
        {
            var bytes = _capture.ToArray();
            if (start < 0 || start > bytes.LongLength)
            {
                throw new ArgumentOutOfRangeException(nameof(start));
            }

            return bytes[(int)start..];
        }
    }

    public override void Flush() => _inner.Flush();

    public override Task FlushAsync(CancellationToken cancellationToken) => _inner.FlushAsync(cancellationToken);

    public override int Read(byte[] buffer, int offset, int count)
    {
        var read = _inner.Read(buffer, offset, count);
        if (!_captureWrites && read > 0)
        {
            Capture(buffer.AsSpan(offset, read));
        }

        return read;
    }

    public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
    {
        var read = await _inner.ReadAsync(buffer, cancellationToken);
        if (!_captureWrites && read > 0)
        {
            Capture(buffer.Span[..read]);
        }

        return read;
    }

    public override void Write(byte[] buffer, int offset, int count)
    {
        if (_captureWrites && count > 0)
        {
            Capture(buffer.AsSpan(offset, count));
        }

        _inner.Write(buffer, offset, count);
    }

    public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
    {
        if (_captureWrites && buffer.Length > 0)
        {
            Capture(buffer.Span);
        }

        await _inner.WriteAsync(buffer, cancellationToken);
    }

    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

    public override void SetLength(long value) => throw new NotSupportedException();

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _capture.Dispose();
            if (!_leaveOpen)
            {
                _inner.Dispose();
            }
        }

        base.Dispose(disposing);
    }

    private void Capture(ReadOnlySpan<byte> buffer)
    {
        lock (_sync)
        {
            _capture.Write(buffer);
        }
    }
}
