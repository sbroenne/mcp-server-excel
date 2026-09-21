using System.Text.Json;
using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.ComInterop.ServiceClient;

/// <summary>
/// Shared serialization for CLI daemon and in-process MCP service contracts.
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

    /// <summary>The service command that produced this response, when available.</summary>
    public string? Command { get; init; }

    /// <summary>The session ID associated with this response, when available.</summary>
    public string? SessionId { get; init; }

    /// <summary>Error message if Success is false.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>Structured error category if Success is false.</summary>
    public string? ErrorCategory { get; init; }

    /// <summary>Exception type that produced the failure, when available.</summary>
    public string? ExceptionType { get; init; }

    /// <summary>HRESULT from a COM failure, when available.</summary>
    [JsonPropertyName("hresult")]
    public string? HResult { get; init; }

    /// <summary>Inner exception details, when available.</summary>
    public string? InnerError { get; init; }

    /// <summary>JSON-serialized result data.</summary>
    public string? Result { get; init; }
}

/// <summary>
/// Service status information.
/// </summary>
public sealed class ServiceStatus
{
    /// <summary>Whether the service is running.</summary>
    public bool Running { get; init; }
    /// <summary>Process ID of the service.</summary>
    public int ProcessId { get; init; }
    /// <summary>Number of active sessions.</summary>
    public int SessionCount { get; init; }
    /// <summary>UTC time when the service started.</summary>
    public DateTime StartTime { get; init; }
    /// <summary>Elapsed time since startup, or zero when stopped.</summary>
    public TimeSpan Uptime => Running ? DateTime.UtcNow - StartTime : TimeSpan.Zero;
}
