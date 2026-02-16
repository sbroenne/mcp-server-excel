using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Spectre.Console;
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
            AnsiConsole.MarkupLine("[red]Session ID is required. Use --session <id>[/]");
            return 1;
        }

        // Action validation
        var rawAction = GetAction(settings);
        if (string.IsNullOrWhiteSpace(rawAction))
        {
            AnsiConsole.MarkupLine("[red]Action is required.[/]");
            return 1;
        }

        var action = rawAction.Trim().ToLowerInvariant();
        if (!ValidActions.Contains(action, StringComparer.OrdinalIgnoreCase))
        {
            var validList = string.Join(", ", ValidActions);
            AnsiConsole.MarkupLine($"[red]Invalid action '{action}'. Valid actions: {validList}[/]");
            return 1;
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
            Console.WriteLine(JsonSerializer.Serialize(
                new { success = false, error = ex.Message },
                ServiceProtocol.JsonOptions));
            return 1;
        }

        // Connect to CLI daemon service (auto-starts if not running)
        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = command,
            SessionId = sessionId,
            Args = args != null ? JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions) : null
        }, cancellationToken);

        // Check for --output file path (generated on all CliSettings)
        var outputPath = settings.GetType().GetProperty("OutputPath")?.GetValue(settings) as string;

        // Output result
        if (response.Success)
        {
            var result = !string.IsNullOrEmpty(response.Result)
                ? response.Result
                : JsonSerializer.Serialize(new { success = true }, ServiceProtocol.JsonOptions);

            if (!string.IsNullOrEmpty(outputPath))
            {
                return WriteOutputToFile(result, outputPath);
            }

            Console.WriteLine(result);
            return 0;
        }
        else
        {
            var errorJson = JsonSerializer.Serialize(
                new { success = false, error = response.ErrorMessage },
                ServiceProtocol.JsonOptions);
            Console.WriteLine(errorJson);
            return 1;
        }
    }

    /// <summary>
    /// Writes the result to a file. For image results containing base64 data,
    /// decodes and writes the binary image. Otherwise writes the JSON text.
    /// </summary>
    private static int WriteOutputToFile(string result, string outputPath)
    {
        try
        {
            // Try to extract base64 image data from the result
            var base64Data = TryExtractBase64Image(result);
            if (base64Data != null)
            {
                var imageBytes = Convert.FromBase64String(base64Data);
                File.WriteAllBytes(outputPath, imageBytes);
                // Write metadata (without the large base64 payload) to stdout
                var doc = JsonDocument.Parse(result);
                var metadata = new Dictionary<string, object?>
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
                Console.WriteLine(JsonSerializer.Serialize(metadata, ServiceProtocol.JsonOptions));
            }
            else
            {
                File.WriteAllText(outputPath, result);
                Console.WriteLine(JsonSerializer.Serialize(
                    new { success = true, outputPath },
                    ServiceProtocol.JsonOptions));
            }
            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine(JsonSerializer.Serialize(
                new { success = false, error = $"Failed to write output: {ex.Message}" },
                ServiceProtocol.JsonOptions));
            return 1;
        }
    }

    /// <summary>
    /// Attempts to extract base64 image data from a JSON result.
    /// Returns the base64 string if found, null otherwise.
    /// </summary>
    private static string? TryExtractBase64Image(string json)
    {
        try
        {
            using var doc = JsonDocument.Parse(json);
            if (doc.RootElement.TryGetProperty("imageBase64", out var imageElement))
            {
                return imageElement.GetString();
            }
        }
        catch (JsonException)
        {
            // Not valid JSON, can't extract image
        }
        return null;
    }
}
