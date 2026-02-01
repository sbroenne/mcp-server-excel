using System.Text.Json;
using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.CLI.Daemon;

/// <summary>
/// Protocol messages for CLI-to-daemon communication over named pipes.
/// Pattern: CLI sends JSON request → Daemon executes → Returns JSON response.
/// All messages are newline-delimited JSON.
/// </summary>
internal static class DaemonProtocol
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
/// Request sent from CLI to daemon.
/// </summary>
internal sealed class DaemonRequest
{
    /// <summary>Command to execute (e.g., "session.open", "sheet.list", "range.get-values").</summary>
    public required string Command { get; init; }

    /// <summary>Session ID for commands that operate on a session.</summary>
    public string? SessionId { get; init; }

    /// <summary>JSON-serialized command arguments.</summary>
    public string? Args { get; init; }
}

/// <summary>
/// Response sent from daemon to CLI.
/// </summary>
internal sealed class DaemonResponse
{
    /// <summary>Whether the command succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message if Success is false.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>JSON-serialized result data.</summary>
    public string? Result { get; init; }
}
