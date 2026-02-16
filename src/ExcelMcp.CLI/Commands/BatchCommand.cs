using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
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
    }

    public override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
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
                Source = "cli-batch"
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
                Error = response.ErrorMessage
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
