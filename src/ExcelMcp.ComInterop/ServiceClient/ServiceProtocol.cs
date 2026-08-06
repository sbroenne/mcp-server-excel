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

    /// <summary>Return a mutation review plan without executing the command.</summary>
    public bool ReviewOnly { get; init; }

    /// <summary>Bound review identifier authorizing a previously reviewed mutation.</summary>
    public string? ReviewId { get; init; }

    /// <summary>Create a recoverable checkpoint immediately before a mutation.</summary>
    public bool Checkpoint { get; init; }

    /// <summary>
    /// Client-generated key that makes an executable session request safe to retry.
    /// An exact retry returns the original receipt; reusing the key for a different
    /// request fails closed.
    /// </summary>
    public string? IdempotencyKey { get; init; }
}

/// <summary>
/// Arguments for one ordered server-side batch request.
/// </summary>
public sealed class ServiceBatchRequest
{
    /// <summary>Operations to execute in their original order.</summary>
    public List<ServiceBatchOperation> Operations { get; init; } = [];

    /// <summary>
    /// Whether execution stops after the first failed operation. The wire-level
    /// default is fail-closed; clients that preserve continue-on-error behavior
    /// must send false explicitly.
    /// </summary>
    public bool StopOnError { get; init; } = true;
}

/// <summary>
/// One operation inside a server-side batch.
/// </summary>
public sealed class ServiceBatchOperation
{
    /// <summary>Command to execute.</summary>
    public required string Command { get; init; }

    /// <summary>Optional session ID, which must match the enclosing request.</summary>
    public string? SessionId { get; init; }

    /// <summary>Command arguments as an embedded JSON value.</summary>
    public JsonElement? Args { get; init; }

    /// <summary>Return a review plan without executing this operation.</summary>
    public bool ReviewOnly { get; init; }

    /// <summary>Review identifier authorizing this operation.</summary>
    public string? ReviewId { get; init; }

    /// <summary>Create a checkpoint before this operation.</summary>
    public bool Checkpoint { get; init; }

    /// <summary>Client-generated retry key for this operation.</summary>
    public string? IdempotencyKey { get; init; }
}

/// <summary>
/// Ordered result envelope returned by a server-side batch.
/// </summary>
public sealed class ServiceBatchResponse
{
    /// <summary>Whether every attempted operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Whether every requested operation was attempted.</summary>
    public bool Completed { get; init; }

    /// <summary>Zero-based index of the first failed operation.</summary>
    public int? FailedIndex { get; init; }

    /// <summary>Results in the exact order in which operations were attempted.</summary>
    public List<ServiceBatchOperationResult> Results { get; init; } = [];
}

/// <summary>
/// Result for one attempted operation in a server-side batch.
/// </summary>
public sealed class ServiceBatchOperationResult
{
    /// <summary>Zero-based position in the request.</summary>
    public int Index { get; init; }

    /// <summary>
    /// Command that was attempted when needed for a standalone validation error.
    /// Normal ordered results omit this redundant copy; callers resolve it from Index.
    /// </summary>
    public string? Command { get; init; }

    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Command result as an embedded JSON value.</summary>
    public JsonElement? Result { get; init; }

    /// <summary>Error message when the operation failed.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>Structured error category.</summary>
    public string? ErrorCategory { get; init; }

    /// <summary>Exception type when available.</summary>
    public string? ExceptionType { get; init; }

    /// <summary>HRESULT from a COM failure, when available.</summary>
    [JsonPropertyName("hresult")]
    public string? HResult { get; init; }
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


