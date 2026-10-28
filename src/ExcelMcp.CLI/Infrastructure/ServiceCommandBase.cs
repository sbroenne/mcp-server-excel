using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.CLI.Telemetry;
using Sbroenne.ExcelMcp.Service;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>
/// Base class for CLI commands that send requests to the service.
/// Handles common validation and execution patterns.
/// </summary>
internal abstract class ServiceCommandBase<TSettings> : AsyncCommand<TSettings>
    where TSettings : CommandSettings
{
    /// <summary>
    /// Gets the session ID from settings.
    /// </summary>
    protected abstract string? GetSessionId(TSettings settings);

    /// <summary>
    /// Gets the action from settings.
    /// </summary>
    protected abstract string? GetAction(TSettings settings);

    /// <summary>
    /// Gets the valid actions for this command.
    /// </summary>
    protected abstract IReadOnlyList<string> ValidActions { get; }

    /// <summary>
    /// Routes the action to a service command and args.
    /// </summary>
    protected abstract (string command, object? args) Route(TSettings settings, string action);

    /// <summary>
    /// Whether this command requires a session ID. Default is true.
    /// Override to return false for commands that don't need a session.
    /// </summary>
    protected virtual bool RequiresSession => true;

    /// <summary>
    /// Validates settings and executes the command.
    /// Returns early with error code if validation fails.
    /// </summary>
    public sealed override async Task<int> ExecuteAsync(CommandContext context, TSettings settings, CancellationToken cancellationToken)
    {
        // Session validation
        var sessionId = GetSessionId(settings);
        if (RequiresSession && string.IsNullOrWhiteSpace(sessionId))
        {
            return CliErrorOutput.WriteError("Session ID is required. Use --session <id>");
        }

        // Action validation
        var rawAction = GetAction(settings);
        if (string.IsNullOrWhiteSpace(rawAction))
        {
            return CliErrorOutput.WriteError("Action is required.");
        }

        var action = rawAction.Trim().ToLowerInvariant();
        if (!ValidActions.Contains(action, StringComparer.OrdinalIgnoreCase))
        {
            var validList = string.Join(", ", ValidActions);
            return CliErrorOutput.WriteError($"Invalid action '{action}'. Valid actions: {validList}");
        }

        // Route and execute
        string command;
        object? args;
        try
        {
            (command, args) = Route(settings, action);
        }
        catch (ArgumentException ex)
        {
            // Parameter validation failed (e.g., required param missing)
            // Return clean JSON error with exit code 1 instead of unhandled crash
            return CliErrorOutput.WriteException(ex, "InvalidInput");
        }

        // Connect to CLI daemon service (auto-starts if not running)
        using var client = await CliCommandRuntime.Current.ClientFactory.ConnectAsync(cancellationToken);
        var request = new ServiceRequest
        {
            Command = command,
            SessionId = sessionId,
            Args = args != null ? JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions) : null
        };
        var response = await CliTelemetry.TrackCommandAsync(
            request,
            () => client.SendAsync(request, cancellationToken));

        // Check for --output file path (generated on all CliSettings)
        var outputPath = settings.GetType().GetProperty("OutputPath")?.GetValue(settings) as string;

        // Output result
        if (response.Success)
        {
            var outcome = CliResultOutput.Read(response, command);
            var originalData = outcome.Data as System.Text.Json.Nodes.JsonObject;
            var durationMs = CliCommandRuntime.Current.OperationTimer.ElapsedMilliseconds;
            if (!outcome.Success)
                return CliResultOutput.Write(outcome, command);

            if (!string.IsNullOrEmpty(outputPath))
            {
                return WriteOutputToFile(outcome, originalData, command, durationMs, outputPath);
            }

            return CliResultOutput.Write(outcome, command);
        }
        else
        {
            return CliResultOutput.Write(response, command);
        }
    }

    /// <summary>
    /// Writes the result to a file. For image results containing base64 data,
    /// decodes and writes the binary image. Otherwise writes the JSON text.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "CLI output-file boundary reports the original write/encoding failure as success=false JSON and returns exit code 1.")]
    private static int WriteOutputToFile(ResultJson outcome, System.Text.Json.Nodes.JsonObject? originalData,
        string command, long durationMs, string outputPath)
    {
        try
        {
            Dictionary<string, object?> metadata;
            var base64Data = originalData?["imageBase64"]?.GetValue<string>();
            if (base64Data != null)
            {
                var imageBytes = Convert.FromBase64String(base64Data);
                File.WriteAllBytes(outputPath, imageBytes);
                // Write metadata (without the large base64 payload) to stdout
                using var doc = JsonDocument.Parse(originalData!.ToJsonString());
                metadata = new Dictionary<string, object?>
                {
                    ["success"] = true,
                    ["outputPath"] = outputPath,
                    ["sizeBytes"] = imageBytes.Length
                };
                if (doc.RootElement.TryGetProperty("width", out var w)) metadata["width"] = w.GetInt32();
                if (doc.RootElement.TryGetProperty("height", out var h)) metadata["height"] = h.GetInt32();
                if (doc.RootElement.TryGetProperty("mimeType", out var m)) metadata["mimeType"] = m.GetString();
                if (doc.RootElement.TryGetProperty("sheetName", out var s)) metadata["sheetName"] = s.GetString();
                if (doc.RootElement.TryGetProperty("rangeAddress", out var r)) metadata["rangeAddress"] = r.GetString();
                if (doc.RootElement.TryGetProperty("message", out var message)) metadata["message"] = message.GetString();
            }
            else
            {
                File.WriteAllText(outputPath, outcome.ToCliJson(command, durationMs));
                metadata = new Dictionary<string, object?>
                {
                    ["success"] = true,
                    ["outputPath"] = outputPath
                };
            }
            metadata["warnings"] = outcome.Warnings;
            metadata["partialState"] = outcome.PartialState;
            return CliResultOutput.Write(ResultJson.Read(
                JsonSerializer.Serialize(metadata, ServiceProtocol.JsonOptions), command), command, durationMs);
        }
        catch (Exception ex)
        {
            return CliResultOutput.Write(outcome.WithFailure(command, ex, "OutputWriteFailed",
                $"The operation completed, but writing its output failed: {ex.Message}",
                partialState: JsonSerializer.Serialize(new
                {
                    operationCompleted = true,
                    outputWriteCompleted = false,
                    operationPartialState = outcome.PartialState
                }, ServiceProtocol.JsonOptions)), command, durationMs);
        }
    }
}
