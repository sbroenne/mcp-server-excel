using System.ComponentModel;
using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Telemetry;
using Sbroenne.ExcelMcp.Core.Utilities;
using Sbroenne.ExcelMcp.Generated;
using Sbroenne.ExcelMcp.Service;
using Spectre.Console.Cli;

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

        [CommandOption("--stream")]
        [Description("Process stdin incrementally: one JSON command per line, one flushed result per command. Waits for more commands until EOF. Cannot be combined with an input file.")]
        public bool Stream { get; init; }
    }

    public override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (settings.Stream)
        {
            return await ExecuteStreamingAsync(settings, cancellationToken);
        }

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

        var validationErrors = new Dictionary<int, string>();
        for (int i = 0; i < commands.Count; i++)
        {
            var cmd = commands[i];
            var argsJson = cmd.Args.HasValue && cmd.Args.Value.ValueKind != JsonValueKind.Undefined
                ? cmd.Args.Value.GetRawText()
                : null;
            var validationStopwatch = Stopwatch.StartNew();
            try
            {
                ServiceRegistry.ValidateCommandArguments(cmd.Command, argsJson);
            }
            catch (Exception ex) when (ex is ArgumentException or JsonException or IOException or UnauthorizedAccessException)
            {
                validationStopwatch.Stop();
                validationErrors[i] = ex.Message;
                CliTelemetry.TrackLocalFailure(
                    cmd.Command,
                    validationStopwatch.ElapsedMilliseconds,
                    "InvalidInput");
            }
        }

        if (validationErrors.Count == commands.Count ||
            (settings.StopOnError && validationErrors.ContainsKey(0)))
        {
            foreach (var validationError in validationErrors.OrderBy(error => error.Key))
            {
                WriteValidationError(validationError.Key, commands[validationError.Key].Command, validationError.Value);
                if (settings.StopOnError)
                {
                    break;
                }
            }
            return 1;
        }

        // Connect to daemon (auto-starts if needed)
        using var client = await CliCommandRuntime.Current.ClientFactory.ConnectAsync(cancellationToken);

        string? activeSession = settings.SessionId;
        bool hasErrors = false;

        for (int i = 0; i < commands.Count; i++)
        {
            var cmd = commands[i];
            if (validationErrors.TryGetValue(i, out var validationError))
            {
                WriteValidationError(i, cmd.Command, validationError);

                hasErrors = true;
                if (settings.StopOnError)
                {
                    break;
                }
                continue;
            }
            var (itemSucceeded, nextSession) = await ExecuteEntryAsync(
                client, cmd, i, activeSession, cancellationToken);
            activeSession = nextSession;

            if (!itemSucceeded)
            {
                hasErrors = true;
                if (settings.StopOnError) break;
            }
        }

        return hasErrors ? 1 : 0;
    }

    private static async Task<int> ExecuteStreamingAsync(Settings settings, CancellationToken cancellationToken)
    {
        if (!string.IsNullOrEmpty(settings.InputFile) && settings.InputFile != "-")
        {
            WriteError("--stream reads NDJSON from stdin; omit --input or use --input -.");
            return 1;
        }

        ICliRequestClient? client = null;
        string? activeSession = settings.SessionId;
        int index = 0;
        bool hasErrors = false;
        try
        {
            while (await CliCommandRuntime.Current.Input.ReadLineAsync(cancellationToken) is { } line)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;

                BatchEntry? entry = null;
                string? validationError = null;
                var validationStopwatch = Stopwatch.StartNew();
                try
                {
                    entry = JsonSerializer.Deserialize<BatchEntry>(line, BatchJsonOptions);
                    if (entry == null || string.IsNullOrWhiteSpace(entry.Command))
                    {
                        throw new ArgumentException("Each line must be a JSON object with a nonempty 'command' field.");
                    }
                    var argsJson = entry.Args.HasValue && entry.Args.Value.ValueKind != JsonValueKind.Undefined
                        ? entry.Args.Value.GetRawText()
                        : null;
                    ServiceRegistry.ValidateCommandArguments(entry.Command, argsJson);
                }
                catch (Exception ex) when (ex is ArgumentException or JsonException or IOException or UnauthorizedAccessException)
                {
                    validationError = ex.Message;
                    var failedCommand = string.IsNullOrWhiteSpace(entry?.Command) ? "batch" : entry.Command;
                    CliTelemetry.TrackLocalFailure(failedCommand, validationStopwatch.ElapsedMilliseconds, "InvalidInput");
                }

                bool itemSucceeded;
                if (validationError != null)
                {
                    WriteValidationError(index, entry?.Command ?? string.Empty, validationError);
                    itemSucceeded = false;
                }
                else
                {
                    // Start/connect only when a valid command arrives, not while waiting on input.
                    client ??= await TryConnectAsync(entry!, index, cancellationToken);
                    if (client == null)
                    {
                        // Reported as this line's result; the next valid line retries the connection.
                        itemSucceeded = false;
                    }
                    else
                    {
                        var outcome = await ExecuteEntryAsync(client, entry!, index, activeSession, cancellationToken);
                        itemSucceeded = outcome.Succeeded;
                        activeSession = outcome.ActiveSession;
                    }
                }

                await CliCommandRuntime.Current.Output.FlushAsync(cancellationToken);
                index++;
                if (!itemSucceeded)
                {
                    hasErrors = true;
                    if (settings.StopOnError) break;
                }
            }
        }
        finally
        {
            client?.Dispose();
        }

        if (index == 0)
        {
            WriteError("No commands provided.");
            return 1;
        }
        return hasErrors ? 1 : 0;
    }

    /// <summary>
    /// Connects for a streamed command. A failure is written as that command's indexed
    /// result and returns null instead of ending the stream; cancellation still propagates.
    /// </summary>
    private static async Task<ICliRequestClient?> TryConnectAsync(
        BatchEntry entry, int index, CancellationToken cancellationToken)
    {
        var stopwatch = Stopwatch.StartNew();
        try
        {
            return await CliCommandRuntime.Current.ClientFactory.ConnectAsync(cancellationToken);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            throw;
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            CliTelemetry.TrackLocalFailure(
                entry.Command,
                stopwatch.ElapsedMilliseconds,
                OperationFailureClassifier.Classify(ex) ?? "ServiceUnavailable");
            CliCommandRuntime.Current.Output.WriteLine(JsonSerializer.Serialize(new BatchResult
            {
                Index = index,
                Command = entry.Command,
                Success = false,
                Error = $"Communication error: {ex.Message}"
            }, BatchJsonOptions));
            return null;
        }
    }

    private static async Task<(bool Succeeded, string? ActiveSession)> ExecuteEntryAsync(
        ICliRequestClient client, BatchEntry cmd, int index, string? activeSession, CancellationToken cancellationToken)
    {
        var sessionId = cmd.SessionId ?? activeSession;
        var request = new ServiceRequest
        {
            Command = cmd.Command,
            SessionId = sessionId,
            Args = cmd.Args.HasValue && cmd.Args.Value.ValueKind != JsonValueKind.Undefined
                ? cmd.Args.Value.GetRawText()
                : null,
            Source = "cli-batch"
        };

        ServiceResponse response;
        try
        {
            response = await CliTelemetry.TrackCommandAsync(
                request, () => client.SendAsync(request, cancellationToken));
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            throw;
        }
        catch (Exception ex)
        {
            response = new ServiceResponse { Success = false, ErrorMessage = $"Communication error: {ex.Message}" };
        }

        var negativeResult = ServiceResultOutcome.TryReadNegative(
            response.Success ? response.Result : null, out var resultErrorMessage, out _);
        var itemSucceeded = response.Success && !negativeResult;
        if (itemSucceeded && activeSession == null &&
            (cmd.Command.Equals("session.open", StringComparison.OrdinalIgnoreCase) ||
             cmd.Command.Equals("session.create", StringComparison.OrdinalIgnoreCase)))
        {
            activeSession = TryExtractSessionId(response.Result);
        }
        if (itemSucceeded && cmd.Command.Equals("session.close", StringComparison.OrdinalIgnoreCase) &&
            string.Equals(sessionId, activeSession, StringComparison.OrdinalIgnoreCase))
        {
            activeSession = null;
        }

        var output = new BatchResult
        {
            Index = index,
            Command = cmd.Command,
            Success = itemSucceeded,
            Result = response.Success ? TryParseJsonElement(response.Result) : null,
            Error = response.ErrorMessage ?? resultErrorMessage ??
                (negativeResult ? "Command reported success: false; see result for details." : null)
        };
        CliCommandRuntime.Current.Output.WriteLine(JsonSerializer.Serialize(output, BatchJsonOptions));
        return (itemSucceeded, activeSession);
    }

    private static void WriteValidationError(int index, string command, string error)
    {
        CliCommandRuntime.Current.Output.WriteLine(JsonSerializer.Serialize(new BatchResult
        {
            Index = index,
            Command = command,
            Success = false,
            Error = error
        }, BatchJsonOptions));
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
            content = await CliCommandRuntime.Current.Input.ReadToEndAsync(cancellationToken);
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
        CliCommandRuntime.Current.Error.WriteLine(
            JsonSerializer.Serialize(new { success = false, error = message }, ServiceProtocol.JsonOptions));
    }

    // JSON options for batch I/O — camelCase, skip nulls for clean output
    private static readonly JsonSerializerOptions BatchJsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = false,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow
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
    }
}
