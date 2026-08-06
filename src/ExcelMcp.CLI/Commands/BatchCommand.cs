using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Service;
using Spectre.Console.Cli;
using ServiceBatchOperation = Sbroenne.ExcelMcp.ComInterop.ServiceClient.ServiceBatchOperation;
using ServiceBatchRequest = Sbroenne.ExcelMcp.ComInterop.ServiceClient.ServiceBatchRequest;
using ServiceBatchResponse = Sbroenne.ExcelMcp.ComInterop.ServiceClient.ServiceBatchResponse;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Executes multiple CLI commands in a single process launch.
/// Reads commands from a JSON file (array) or stdin (NDJSON), sends each
/// to the daemon sequentially, and outputs NDJSON results (one per line).
///
/// Session auto-capture: if a session.open/create succeeds and no --session
/// was provided, the returned sessionId becomes the default for subsequent commands.
/// </summary>
internal sealed class BatchCommand : AsyncCommand<BatchCommand.Settings>
{
    internal const int MaximumServerBatchOperations = 256;
    internal sealed class Settings : CommandSettings
    {
        [CommandOption("-i|--input <FILE>")]
        [Description("JSON file with command array. Use '-' for stdin (NDJSON, one command per line). If omitted, reads from stdin.")]
        public string? InputFile { get; init; }

        [CommandOption("-s|--session <SESSION>")]
        [Description("Default session ID for all commands. Overridden by per-command sessionId. Auto-captured from session.open/create if not set.")]
        public string? SessionId { get; init; }

        [CommandOption("--stop-on-error")]
        [Description("Stop execution on first error (default: continue all commands).")]
        public bool StopOnError { get; init; }
    }

    protected override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        // Read commands from file or stdin
        List<BatchEntry> commands;
        try
        {
            commands = await ReadCommandsAsync(settings.InputFile, cancellationToken);
        }
        catch (Exception ex)
        {
            WriteError($"Failed to read commands: {ex.Message}");
            return 1;
        }

        if (commands.Count == 0)
        {
            WriteError("No commands provided.");
            return 1;
        }

        // Validate all commands have a command field
        for (int i = 0; i < commands.Count; i++)
        {
            if (string.IsNullOrWhiteSpace(commands[i].Command))
            {
                WriteError($"Command at index {i} is missing the 'command' field.");
                return 1;
            }
        }

        // Connect to daemon (auto-starts if needed)
        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);

        string? activeSession = settings.SessionId;
        if (CanUseServerBatch(commands, activeSession))
        {
            int? serverBatchExitCode = await TryExecuteServerBatchAsync(
                client,
                commands,
                activeSession!,
                settings.StopOnError,
                cancellationToken);
            if (serverBatchExitCode.HasValue)
            {
                return serverBatchExitCode.Value;
            }
        }

        bool hasErrors = false;

        for (int i = 0; i < commands.Count; i++)
        {
            var cmd = commands[i];
            var sessionId = cmd.SessionId ?? activeSession;

            // Build the service request
            var request = new ServiceRequest
            {
                Command = cmd.Command,
                SessionId = sessionId,
                Args = cmd.Args.HasValue && cmd.Args.Value.ValueKind != JsonValueKind.Undefined
                    ? cmd.Args.Value.GetRawText()
                    : null,
                Source = "cli-batch",
                ReviewOnly = cmd.ReviewOnly,
                ReviewId = cmd.ReviewId,
                Checkpoint = cmd.Checkpoint,
                IdempotencyKey = cmd.IdempotencyKey
            };

            ServiceResponse response;
            try
            {
                response = await client.SendAsync(request, cancellationToken);
            }
            catch (Exception ex)
            {
                response = new ServiceResponse { Success = false, ErrorMessage = $"Communication error: {ex.Message}" };
            }

            // Auto-capture sessionId from session.open/create results
            if (response.Success && activeSession == null &&
                (cmd.Command.Equals("session.open", StringComparison.OrdinalIgnoreCase) ||
                 cmd.Command.Equals("session.create", StringComparison.OrdinalIgnoreCase)))
            {
                activeSession = TryExtractSessionId(response.Result);
            }

            // Auto-clear session on session.close
            if (response.Success &&
                cmd.Command.Equals("session.close", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(sessionId, activeSession, StringComparison.OrdinalIgnoreCase))
            {
                activeSession = null;
            }

            // Output result as NDJSON line
            var output = new BatchResult
            {
                Index = i,
                Command = cmd.Command,
                Success = response.Success,
                Result = response.Success ? TryParseJsonElement(response.Result) : null,
                Error = response.ErrorMessage,
                ErrorCategory = response.ErrorCategory,
                ExceptionType = response.ExceptionType,
                HResult = response.HResult,
            };

            Console.WriteLine(JsonSerializer.Serialize(output, BatchJsonOptions));

            if (!response.Success)
            {
                hasErrors = true;
                if (settings.StopOnError) break;
            }
        }

        return hasErrors ? 1 : 0;
    }

    private static bool CanUseServerBatch(IReadOnlyList<BatchEntry> commands, string? activeSession)
    {
        if (string.IsNullOrWhiteSpace(activeSession) ||
            commands.Count is < 1 or > MaximumServerBatchOperations)
        {
            return false;
        }

        foreach (var command in commands)
        {
            var effectiveSession = command.SessionId ?? activeSession;
            if (!string.Equals(effectiveSession, activeSession, StringComparison.Ordinal))
            {
                return false;
            }

            var parts = command.Command.Split('.', 2);
            if (parts.Length != 2 || !IsServerBatchCategory(parts[0], parts[1]))
            {
                return false;
            }
        }

        return true;
    }

    private static bool IsServerBatchCategory(string category, string action) =>
        ServerBatchCategories.Contains(category) ||
        (category == "sheet" && action is not ("copy-to-file" or "move-to-file"));

    internal static bool IsWithinServerBatchLimit(int commandCount) =>
        commandCount is >= 1 and <= MaximumServerBatchOperations;

    private static async Task<int?> TryExecuteServerBatchAsync(
        ServiceClient client,
        IReadOnlyList<BatchEntry> commands,
        string sessionId,
        bool stopOnError,
        CancellationToken cancellationToken)
    {
        var batchRequest = new ServiceBatchRequest
        {
            StopOnError = stopOnError,
            Operations = commands.Select(command => new ServiceBatchOperation
            {
                Command = command.Command,
                SessionId = command.SessionId ?? sessionId,
                Args = command.Args,
                ReviewOnly = command.ReviewOnly,
                ReviewId = command.ReviewId,
                Checkpoint = command.Checkpoint,
                IdempotencyKey = command.IdempotencyKey,
            }).ToList(),
        };
        var request = new ServiceRequest
        {
            Command = "session.batch",
            SessionId = sessionId,
            Args = ServiceProtocol.Serialize(batchRequest),
            Source = "cli-batch",
        };

        ServiceResponse response;
        try
        {
            response = await client.SendAsync(request, cancellationToken);
        }
        catch (Exception ex)
        {
            WriteBatchTransportFailure(ex.Message);
            return 1;
        }

        // An older daemon has definitely rejected the batch before executing any
        // nested command, so falling back is safe and preserves compatibility.
        if (!response.Success &&
            string.Equals(response.ErrorMessage, "Unknown session action: batch", StringComparison.Ordinal))
        {
            return null;
        }

        ServiceBatchResponse? batchResponse = null;
        if (!string.IsNullOrWhiteSpace(response.Result))
        {
            try
            {
                batchResponse = ServiceProtocol.Deserialize<ServiceBatchResponse>(response.Result);
            }
            catch (JsonException)
            {
                // An ambiguous or malformed response must never trigger a replay.
            }
        }

        if (!TryValidateServerBatchResponse(batchResponse, commands.Count, out var protocolError) || batchResponse is null)
        {
            WriteBatchTransportFailure(protocolError ?? response.ErrorMessage ?? "The service returned an invalid batch response.");
            return 1;
        }

        foreach (var result in batchResponse.Results)
        {
            string command = result.Command ??
                (result.Index >= 0 && result.Index < commands.Count
                    ? commands[result.Index].Command
                    : "session.batch");
            var output = new BatchResult
            {
                Index = result.Index,
                Command = command,
                Success = result.Success,
                Result = result.Result,
                Error = result.ErrorMessage,
                ErrorCategory = result.ErrorCategory,
                ExceptionType = result.ExceptionType,
                HResult = result.HResult,
            };
            Console.WriteLine(JsonSerializer.Serialize(output, BatchJsonOptions));
        }

        return batchResponse.Success ? 0 : 1;
    }

    internal static bool TryValidateServerBatchResponse(
        ServiceBatchResponse? batchResponse,
        int commandCount,
        out string? error)
    {
        error = null;
        if (batchResponse is null)
        {
            error = "The service returned an invalid batch response.";
            return false;
        }

        var results = batchResponse.Results;
        if (results is null || results.Count is < 1 || results.Count > commandCount)
        {
            error = "The service returned an invalid batch response result count.";
            return false;
        }

        // A pre-execution validation failure may be represented by a single
        // result at the original operation position (for example, operation
        // 3 can be rejected before operations 0-2 are dispatched).  Normal
        // execution results are always contiguous from index 0.
        int expectedStart = !batchResponse.Completed &&
            batchResponse.FailedIndex is { } validationIndex &&
            results.Count == 1 && results[0].Index == validationIndex
            ? validationIndex
            : 0;

        for (int index = 0; index < results.Count; index++)
        {
            var result = results[index];
            if (result is null || result.Index != expectedStart + index ||
                result.Index < 0 || result.Index >= commandCount)
            {
                error = "The service returned non-sequential or out-of-range batch result indexes.";
                return false;
            }
        }

        if (batchResponse.Completed && results.Count != commandCount)
        {
            error = "The service marked the batch complete without returning every operation result.";
            return false;
        }

        if (batchResponse.Success &&
            (!batchResponse.Completed || results.Any(result => !result.Success)))
        {
            error = "The service returned a successful batch envelope with incomplete or failed operation results.";
            return false;
        }

        if (!batchResponse.Success && results.All(result => result.Success))
        {
            error = "The service returned a failed batch envelope without a failed operation result.";
            return false;
        }

        if (batchResponse.FailedIndex is { } failedIndex &&
            (failedIndex < 0 || failedIndex >= commandCount ||
             !results.Any(result => result.Index == failedIndex && !result.Success)))
        {
            error = "The service returned an invalid failed operation index.";
            return false;
        }

        return true;
    }

    private static void WriteBatchTransportFailure(string message)
    {
        var output = new BatchResult
        {
            Index = 0,
            Command = "session.batch",
            Success = false,
            Error = $"Communication error: {message}",
            ErrorCategory = "ServiceProtocol",
        };
        Console.WriteLine(JsonSerializer.Serialize(output, BatchJsonOptions));
    }

    /// <summary>
    /// Reads commands from a JSON file (array format) or stdin (NDJSON format).
    /// Auto-detects format: if content starts with '[', parses as JSON array; otherwise NDJSON.
    /// </summary>
    private static async Task<List<BatchEntry>> ReadCommandsAsync(string? inputFile, CancellationToken cancellationToken)
    {
        string content;

        if (string.IsNullOrEmpty(inputFile) || inputFile == "-")
        {
            // Read from stdin
            content = await Console.In.ReadToEndAsync(cancellationToken);
        }
        else
        {
            // Read from file
            var fullPath = Path.GetFullPath(inputFile);
            if (!File.Exists(fullPath))
            {
                throw new FileNotFoundException($"Input file not found: {fullPath}");
            }
            content = await File.ReadAllTextAsync(fullPath, cancellationToken);
        }

        content = content.Trim();

        if (string.IsNullOrEmpty(content))
        {
            return [];
        }

        // Auto-detect format: JSON array vs NDJSON
        if (content.StartsWith('['))
        {
            // JSON array format
            return JsonSerializer.Deserialize<List<BatchEntry>>(content, BatchJsonOptions) ?? [];
        }

        // NDJSON format: one JSON object per non-empty line
        var commands = new List<BatchEntry>();
        foreach (var line in content.Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            var trimmed = line.Trim();
            if (string.IsNullOrEmpty(trimmed)) continue;

            var entry = JsonSerializer.Deserialize<BatchEntry>(trimmed, BatchJsonOptions);
            if (entry != null)
            {
                commands.Add(entry);
            }
        }

        return commands;
    }

    /// <summary>
    /// Extracts sessionId from a session.open/create result JSON string.
    /// </summary>
    private static string? TryExtractSessionId(string? resultJson)
    {
        if (string.IsNullOrEmpty(resultJson)) return null;

        try
        {
            using var doc = JsonDocument.Parse(resultJson);
            if (doc.RootElement.TryGetProperty("sessionId", out var sessionIdProp) &&
                sessionIdProp.ValueKind == JsonValueKind.String)
            {
                return sessionIdProp.GetString();
            }
        }
        catch (JsonException)
        {
            // Not valid JSON — ignore
        }

        return null;
    }

    /// <summary>
    /// Parses a JSON string into a JsonElement for embedding in the output.
    /// </summary>
    private static JsonElement? TryParseJsonElement(string? json)
    {
        if (string.IsNullOrEmpty(json)) return null;

        try
        {
            using var doc = JsonDocument.Parse(json);
            return doc.RootElement.Clone();
        }
        catch (JsonException)
        {
            return null;
        }
    }

    private static void WriteError(string message)
    {
        Console.Error.WriteLine(JsonSerializer.Serialize(new { success = false, error = message }, ServiceProtocol.JsonOptions));
    }

    // JSON options for batch I/O — camelCase, skip nulls for clean output
    private static readonly JsonSerializerOptions BatchJsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private static readonly HashSet<string> ServerBatchCategories = new(StringComparer.Ordinal)
    {
        "sheetstyle",
        "range",
        "rangeedit",
        "rangeformat",
        "rangelink",
        "table",
        "tablecolumn",
        "powerquery",
        "pivottable",
        "pivottablefield",
        "pivottablecalc",
        "chart",
        "chartconfig",
        "connection",
        "calculation",
        "namedrange",
        "conditionalformat",
        "vba",
        "datamodel",
        "datamodelrel",
        "slicer",
        "screenshot",
        "window",
        "pythoninexcel",
    };

    // ── Models ──────────────────────────────────────────────────────

    private sealed class BatchEntry
    {
        [JsonPropertyName("command")]
        public string Command { get; init; } = string.Empty;

        [JsonPropertyName("sessionId")]
        public string? SessionId { get; init; }

        [JsonPropertyName("args")]
        public JsonElement? Args { get; init; }

        [JsonPropertyName("reviewOnly")]
        public bool ReviewOnly { get; init; }

        [JsonPropertyName("reviewId")]
        public string? ReviewId { get; init; }

        [JsonPropertyName("checkpoint")]
        public bool Checkpoint { get; init; }

        [JsonPropertyName("idempotencyKey")]
        public string? IdempotencyKey { get; init; }
    }

    private sealed class BatchResult
    {
        [JsonPropertyName("index")]
        public int Index { get; init; }

        [JsonPropertyName("command")]
        public string Command { get; init; } = string.Empty;

        [JsonPropertyName("success")]
        public bool Success { get; init; }

        [JsonPropertyName("result")]
        public JsonElement? Result { get; init; }

        [JsonPropertyName("error")]
        public string? Error { get; init; }

        [JsonPropertyName("errorCategory")]
        public string? ErrorCategory { get; init; }

        [JsonPropertyName("exceptionType")]
        public string? ExceptionType { get; init; }

        [JsonPropertyName("hresult")]
        public string? HResult { get; init; }
    }
}
