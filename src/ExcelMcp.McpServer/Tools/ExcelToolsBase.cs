using System.Diagnostics.CodeAnalysis;
using System.Text.Json;
using System.Text.Json.Serialization;
using ModelContextProtocol;
using Sbroenne.ExcelMcp.ComInterop.Session;

#pragma warning disable IL2070 // 'this' argument does not satisfy 'DynamicallyAccessedMembersAttribute' requirements

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Base class for Excel MCP tools providing common patterns and utilities.
/// All Excel tools inherit from this to ensure consistency for LLM usage.
/// Provides session management support for conversational workflow performance.
/// </summary>
[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicMethods)]
public static class ExcelToolsBase
{
    private static readonly SessionManager SessionManager = new();

    /// <summary>
    /// JSON serializer options with enum string conversion for user-friendly API responses.
    /// Used by all Excel tools for consistent JSON formatting.
    /// </summary>
    public static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Converters = { new JsonStringEnumConverter() }
    };

    /// <summary>
    /// Gets the SessionManager instance for session lifecycle operations.
    /// </summary>
    public static SessionManager GetSessionManager() => SessionManager;

    /// <summary>
    /// Executes an async Core command with session management.
    /// Uses the provided sessionId to retrieve an active session from SessionManager.
    /// This is the new session-based pattern that replaces batch-of-one operations.
    /// </summary>
    /// <typeparam name="T">Return type of the command</typeparam>
    /// <param name="sessionId">Required session ID from excel_file 'open' action</param>
    /// <param name="action">Async action that takes IExcelBatch and returns Task&lt;T&gt;</param>
    /// <returns>Result of the command</returns>
    /// <exception cref="McpException">Session not found or command execution failed</exception>
    public static async Task<T> WithSessionAsync<T>(
        string sessionId,
        Func<IExcelBatch, Task<T>> action)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            throw new McpException("sessionId is required. Use excel_file 'open' action to start a session.");
        }

        var batch = SessionManager.GetSession(sessionId);
        if (batch == null)
        {
            var activeSessionIds = SessionManager.ActiveSessionIds.ToList();
            var sessionCount = activeSessionIds.Count;
            var errorMessage = sessionCount switch
            {
                0 => $"Session '{sessionId}' not found. No active sessions exist.",
                1 => $"Session '{sessionId}' not found. Active session: {activeSessionIds[0]}",
                _ => $"Session '{sessionId}' not found. {sessionCount} active sessions exist."
            };
            throw new McpException(errorMessage);
        }

        return await action(batch);
    }

    /// <summary>
    /// Throws MCP exception for unknown actions.
    /// SDK Pattern: Use McpException for parameter validation errors.
    /// </summary>
    /// <param name="action">The invalid action that was attempted</param>
    /// <param name="supportedActions">List of supported actions for this tool</param>
    /// <exception cref="McpException">Always throws with descriptive error message</exception>
    public static void ThrowUnknownAction(string action, params string[] supportedActions)
    {
        throw new McpException(
            $"Unknown action '{action}'. Supported: {string.Join(", ", supportedActions)}");
    }

    /// <summary>
    /// Throws MCP exception for missing required parameters.
    /// SDK Pattern: Use McpException for parameter validation errors.
    /// </summary>
    /// <param name="parameterName">Name of the missing parameter</param>
    /// <param name="action">The action that requires the parameter</param>
    /// <exception cref="McpException">Always throws with descriptive error message</exception>
    public static void ThrowMissingParameter(string parameterName, string action)
    {
        throw new McpException(
            $"{parameterName} is required for {action} action");
    }

    /// <summary>
    /// Wraps exceptions in MCP exceptions for better error reporting.
    /// SDK Pattern: Wrap business logic exceptions in McpException with context.
    /// LLM-Optimized: Include full exception details including stack trace context for debugging.
    /// </summary>
    /// <param name="ex">The exception that occurred</param>
    /// <param name="action">The action that was being attempted</param>
    /// <param name="filePath">The file path involved (optional)</param>
    /// <exception cref="McpException">Always throws with contextual error message</exception>
    public static void ThrowInternalError(Exception ex, string action, string? filePath = null)
    {
        // Build comprehensive error message for LLM debugging
        var errorMessage = filePath != null
            ? $"{action} failed for '{filePath}': {ex.Message}"
            : $"{action} failed: {ex.Message}";

        // Include exception type and inner exception details for better diagnostics
        if (ex.InnerException != null)
        {
            errorMessage += $" (Inner: {ex.InnerException.Message})";
        }

        // Add exception type to help identify the root cause
        errorMessage += $" [Exception Type: {ex.GetType().Name}]";

        throw new McpException(errorMessage, ex);
    }

    /// <summary>
    /// Converts Pascal/camelCase text to kebab-case for consistent naming.
    /// Used internally for action parameter normalization.
    /// </summary>
    /// <param name="text">Text to convert</param>
    /// <returns>kebab-case version of the text</returns>
    public static string ToKebabCase(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        var result = new System.Text.StringBuilder();
        for (int i = 0; i < text.Length; i++)
        {
            if (i > 0 && char.IsUpper(text[i]))
            {
                result.Append('-');
            }
            result.Append(char.ToLowerInvariant(text[i]));
        }
        return result.ToString();
    }
}
