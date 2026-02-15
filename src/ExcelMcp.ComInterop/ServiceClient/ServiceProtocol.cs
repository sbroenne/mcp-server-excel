using System.Text.Json;
using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.ComInterop.ServiceClient;

/// <summary>
/// Protocol messages for CLI/MCP-to-service communication over named pipes.
/// Pattern: Client sends JSON request → Service executes → Returns JSON response.
/// All messages are newline-delimited JSON.
/// </summary>
public static class ServiceProtocol
{
    /// <summary>
    /// JSON serializer options for service protocol messages.
    /// </summary>
    public static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = false,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter() }
    };

    /// <summary>
    /// Serializes a message to JSON.
    /// </summary>
    public static string Serialize<T>(T message) => JsonSerializer.Serialize(message, JsonOptions);

    /// <summary>
    /// Deserializes a message from JSON.
    /// </summary>
    public static T? Deserialize<T>(string json) => JsonSerializer.Deserialize<T>(json, JsonOptions);
}

/// <summary>
/// Request sent from client (CLI or MCP) to service.
/// </summary>
public sealed class ServiceRequest
{
    /// <summary>Command to execute (e.g., "session.open", "sheet.list", "range.get-values").</summary>
    public required string Command { get; init; }

    /// <summary>Session ID for commands that operate on a session.</summary>
    public string? SessionId { get; init; }

    /// <summary>JSON-serialized command arguments.</summary>
    public string? Args { get; init; }

    /// <summary>Source of the request (CLI or MCP).</summary>
    public string? Source { get; init; }
}

/// <summary>
/// Response sent from service to client.
/// </summary>
public sealed class ServiceResponse
{
    /// <summary>Whether the command succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message if Success is false.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>JSON-serialized result data.</summary>
    public string? Result { get; init; }
}


