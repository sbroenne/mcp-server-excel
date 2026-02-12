using System.Text.Json;

// Re-export shared types from ComInterop for CLI internal use
using SharedProtocol = Sbroenne.ExcelMcp.ComInterop.ServiceClient.ServiceProtocol;
using SharedRequest = Sbroenne.ExcelMcp.ComInterop.ServiceClient.ServiceRequest;
using SharedResponse = Sbroenne.ExcelMcp.ComInterop.ServiceClient.ServiceResponse;

namespace Sbroenne.ExcelMcp.Service;

/// <summary>
/// Protocol messages for CLI-to-service communication over named pipes.
/// Pattern: CLI sends JSON request → Service executes → Returns JSON response.
/// All messages are newline-delimited JSON.
///
/// NOTE: This is a thin wrapper over the shared ComInterop.ServiceClient types
/// for backward compatibility within CLI.
/// </summary>
public static class ServiceProtocol
{
    public static readonly JsonSerializerOptions JsonOptions = SharedProtocol.JsonOptions;

    public static string Serialize<T>(T message) => SharedProtocol.Serialize(message);
    public static T? Deserialize<T>(string json) => SharedProtocol.Deserialize<T>(json);
}

/// <summary>
/// Request sent from CLI to service.
/// </summary>
public sealed class ServiceRequest
{
    /// <summary>Command to execute (e.g., "session.open", "sheet.list", "range.get-values").</summary>
    public required string Command { get; init; }

    /// <summary>Session ID for commands that operate on a session.</summary>
    public string? SessionId { get; init; }

    /// <summary>JSON-serialized command arguments.</summary>
    public string? Args { get; init; }

    /// <summary>Source of the request.</summary>
    public string? Source { get; init; }

    /// <summary>
    /// Converts to shared request type for serialization.
    /// </summary>
    public SharedRequest ToShared() => new()
    {
        Command = Command,
        SessionId = SessionId,
        Args = Args,
        Source = Source ?? "cli"
    };
}

/// <summary>
/// Response sent from service to CLI.
/// </summary>
public sealed class ServiceResponse
{
    /// <summary>Whether the command succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message if Success is false.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>JSON-serialized result data.</summary>
    public string? Result { get; init; }

    /// <summary>
    /// Creates from shared response type.
    /// </summary>
    public static ServiceResponse FromShared(SharedResponse shared) => new()
    {
        Success = shared.Success,
        ErrorMessage = shared.ErrorMessage,
        Result = shared.Result
    };
}


