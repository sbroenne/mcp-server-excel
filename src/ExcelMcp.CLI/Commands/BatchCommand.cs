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
        [Description("JSON array or NDJSON (one command per line), from a file or stdin. Use '-' or omit --input for stdin.")]
        public string? InputFile { get; init; }

        [CommandOption("-s|--session <SESSION>")]
        [Description("Default session ID. Overridden by per-command sessionId. Captured from session.open/create/attach if not set.")]
        public string? SessionId { get; init; }

        [CommandOption("--stop-on-error")]
        [Description("Stop execution on first error (default: continue all commands).")]
        public bool StopOnError { get; init; }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "CLI batch boundary writes explicit failed NDJSON items or input errors and exits nonzero; communication/validation exceptions are surfaced with their original messages, not converted into successful operations.")]
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
            CliCommandRuntime.Current.Error.WriteLine(BatchHelp.InputHint);
            return CliErrorOutput.WriteException(ex);
        }

        if (commands.Count == 0)
        {
            WriteError($"No commands provided. {BatchHelp.InputHint}");
            return 1;
        }

        // Validate all commands have a command field
        for (int i = 0; i < commands.Count; i++)
        {
            if (string.IsNullOrWhiteSpace(commands[i].Command))
            {
                WriteError($"Command at index {i} is missing the 'command' field. {BatchHelp.InputHint}");
                return 1;
            }
        }

        var validationErrors = new Dictionary<int, (Exception Cause, long DurationMs)>();
        var validationDurations = new long[commands.Count];
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
                validationErrors[i] = (ex, validationStopwatch.ElapsedMilliseconds);
                CliTelemetry.TrackLocalFailure(
                    cmd.Command,
                    validationStopwatch.ElapsedMilliseconds,
                    "InvalidInput");
            }
            validationDurations[i] = validationStopwatch.ElapsedMilliseconds;
        }

        if (validationErrors.Count == commands.Count ||
            (settings.StopOnError && validationErrors.ContainsKey(0)))
        {
            foreach (var validationError in validationErrors.OrderBy(error => error.Key))
            {
                WriteValidationError(validationError.Key, commands[validationError.Key],
                    validationError.Value.Cause, validationError.Value.DurationMs);
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
            var itemTimer = Stopwatch.StartNew();
            var sessionId = cmd.SessionId ?? activeSession;
            if (validationErrors.TryGetValue(i, out var validationError))
            {
                WriteValidationError(i, cmd, validationError.Cause, validationError.DurationMs);

                hasErrors = true;
                if (settings.StopOnError)
                {
                    break;
                }
                continue;
            }
            var argsJson = cmd.Args.HasValue && cmd.Args.Value.ValueKind != JsonValueKind.Undefined
                ? cmd.Args.Value.GetRawText()
                : null;
            // Build the service request
            var request = new ServiceRequest
            {
                Command = cmd.Command,
                SessionId = sessionId,
                Args = argsJson,
                Source = "cli-batch"
            };

            ServiceResponse response;
            CliCommandRuntime.Current.Error.WriteLine($"Batch {i + 1}/{commands.Count}: {cmd.Command} started.");
            await CliCommandRuntime.Current.Error.FlushAsync(CancellationToken.None);
            try
            {
                response = await CliTelemetry.TrackCommandAsync(
                    request,
                    () => client.SendAsync(request, cancellationToken));
            }
            catch (Exception ex)
            {
                response = new ServiceResponse
                {
                    Success = false,
                    ErrorMessage = $"Communication error: {ex.Message}",
                    ErrorCategory = OperationFailureClassifier.Classify(ex),
                    ExceptionType = ex.GetType().Name,
                    HResult = OperationFailureClassifier.GetComHResult(ex),
                    InnerError = ex.InnerException?.Message
                };
            }

            var outcome = CliResultOutput.Read(response, cmd.Command);
            // Auto-capture sessionId from session.open/create/attach results
            if (outcome.Success && activeSession == null &&
                (cmd.Command.Equals("session.open", StringComparison.OrdinalIgnoreCase) ||
                 cmd.Command.Equals("session.attach", StringComparison.OrdinalIgnoreCase) ||
                 cmd.Command.Equals("session.create", StringComparison.OrdinalIgnoreCase)))
            {
                activeSession = outcome.Data?["sessionId"]?.GetValue<string>();
            }

            // Auto-clear session on session.close
            if (outcome.Success &&
                cmd.Command.Equals("session.close", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(sessionId, activeSession, StringComparison.OrdinalIgnoreCase))
            {
                activeSession = null;
            }

            // Output result as NDJSON line. A delivered result reporting success:false
            // is a failed item, matching single-command exit codes and MCP isError.
            var itemSucceeded = outcome.Success;
            CliResultOutput.Write(outcome, cmd.Command, itemTimer.ElapsedMilliseconds + validationDurations[i], i);
            await CliCommandRuntime.Current.Output.FlushAsync(CancellationToken.None);
            CliCommandRuntime.Current.Error.WriteLine(
                $"Batch {i + 1}/{commands.Count}: {cmd.Command} {(itemSucceeded ? "completed" : "failed")}.");

            if (!itemSucceeded)
            {
                hasErrors = true;
                if (settings.StopOnError || cancellationToken.IsCancellationRequested ||
                    outcome.Error?["category"]?.GetValue<string>() == "Cancelled") break;
            }
        }

        return hasErrors ? 1 : 0;
    }

    private static string ExplainValidationError(string command, string error)
    {
        var parts = command.Split('.', 2);
        if (parts.Length == 2 &&
            _CliCategoryMetadata.ActionParametersByCommand.TryGetValue(parts[0], out var actions) &&
            actions.ContainsKey(parts[1]))
        {
            return $"{error} See 'excelcli {parts[0]} {parts[1]} --help' for batch argument names.";
        }

        var words = command.Split([' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
        if (words.Length == 2 &&
            _CliCategoryMetadata.ActionParametersByCommand.TryGetValue(words[0], out var validActions))
        {
            var action = validActions.Keys.FirstOrDefault(candidate =>
                string.Equals(candidate, words[1], StringComparison.OrdinalIgnoreCase));
            if (action != null)
            {
                var group = _CliCategoryMetadata.Categories.Single(category =>
                    string.Equals(category.CliCommandName, words[0], StringComparison.OrdinalIgnoreCase));
                return $"{error} Use '{group.CliCommandName}.{action}' for the command field. {BatchHelp.InputHint}";
            }
        }

        return $"{error} Batch commands use 'group.action', for example 'diag.ping'. {BatchHelp.InputHint}";
    }

    private static void WriteValidationError(int index, BatchEntry entry, Exception error, long durationMs)
    {
        CliResultOutput.Write(new ServiceResponse
        {
            Command = entry.Command,
            Success = false,
            ErrorCategory = "InvalidInput",
            ErrorMessage = ExplainValidationError(entry.Command, error.Message),
            ExceptionType = error.GetType().Name,
            HResult = OperationFailureClassifier.GetComHResult(error),
            InnerError = error.InnerException?.Message
        }, entry.Command, durationMs: durationMs, index: index);
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
            var entries = JsonSerializer.Deserialize<List<BatchEntry>>(content, BatchJsonOptions)
                ?? throw new ArgumentException("Batch input must contain command objects.");
            if (entries.Any(entry => entry is null))
                throw new ArgumentException("Every batch item must be a command object, not null.");
            return entries;
        }

        // NDJSON format: one JSON object per non-empty line
        var commands = new List<BatchEntry>();
        foreach (var line in content.Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            var trimmed = line.Trim();
            if (string.IsNullOrEmpty(trimmed)) continue;

            var entry = JsonSerializer.Deserialize<BatchEntry>(trimmed, BatchJsonOptions);
            commands.Add(entry ?? throw new ArgumentException("Every batch line must be a command object, not null."));
        }

        return commands;
    }

    private static void WriteError(string message)
    {
        CliErrorOutput.WriteError(message, "InvalidInput");
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

}
