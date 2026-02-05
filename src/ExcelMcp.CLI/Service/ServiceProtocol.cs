using System.Text.Json;
using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.CLI.Service;

/// <summary>
/// Protocol messages for CLI-to-service communication over named pipes.
/// Pattern: CLI sends JSON request → Service executes → Returns JSON response.
/// All messages are newline-delimited JSON.
/// </summary>
internal static class ServiceProtocol
{
    public static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = false,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter() }
    };

    public static string Serialize<T>(T message) => JsonSerializer.Serialize(message, JsonOptions);
    public static T? Deserialize<T>(string json) => JsonSerializer.Deserialize<T>(json, JsonOptions);
}

/// <summary>
/// Request sent from CLI to service.
/// </summary>
internal sealed class ServiceRequest
{
    /// <summary>Command to execute (e.g., "session.open", "sheet.list", "range.get-values").</summary>
    public required string Command { get; init; }

    /// <summary>Session ID for commands that operate on a session.</summary>
    public string? SessionId { get; init; }

    /// <summary>JSON-serialized command arguments.</summary>
    public string? Args { get; init; }
}

/// <summary>
/// Response sent from service to CLI.
/// </summary>
internal sealed class ServiceResponse
{
    /// <summary>Whether the command succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message if Success is false.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>JSON-serialized result data.</summary>
    public string? Result { get; init; }
}
