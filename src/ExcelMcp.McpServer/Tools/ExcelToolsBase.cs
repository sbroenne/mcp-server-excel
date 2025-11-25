using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Text.Json;
using System.Text.Json.Serialization;
using ModelContextProtocol;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.McpServer.Telemetry;

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
    /// Executes a synchronous Core command with session management.
    /// Uses the provided sessionId to retrieve an active session from SessionManager.
    /// All Core commands are now synchronous (blocking).
    /// </summary>
    /// <typeparam name="T">Return type of the command</typeparam>
    /// <param name="sessionId">Required session ID from excel_file 'open' action</param>
    /// <param name="action">Synchronous action that takes IExcelBatch and returns T</param>
    /// <returns>Result of the command</returns>
    /// <exception cref="McpException">Session not found or command execution failed</exception>
    public static T WithSession<T>(
        string sessionId,
        Func<IExcelBatch, T> action)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            throw new ArgumentException("sessionId is required. Use excel_file 'open' action to start a session.", nameof(sessionId));
        }

        var batch = SessionManager.GetSession(sessionId);
        if (batch == null)
        {
            var activeSessionIds = SessionManager.ActiveSessionIds.ToList();
            var sessionCount = activeSessionIds.Count;
            var errorMessage = sessionCount switch
            {
                0 => $"Session '{sessionId}' not found. No active sessions exist. " +
                     "Possible causes: (1) Session was closed prematurely before completing operations, " +
                     "(2) Session never created. " +
                     "Recovery: Use excel_file(action='open') to create a new session.",
                1 => $"Session '{sessionId}' not found. Active session: {activeSessionIds[0]}. " +
                     "Did you close the session before completing all operations? Use the active sessionId shown above.",
                _ => $"Session '{sessionId}' not found. {sessionCount} active sessions exist. " +
                     "Verify you're using the correct sessionId from excel_file 'open' action."
            };
            throw new InvalidOperationException(errorMessage);
        }

        return action(batch);
    }

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
            throw new ArgumentException("sessionId is required. Use excel_file 'open' action to start a session.", nameof(sessionId));
        }

        var batch = SessionManager.GetSession(sessionId);
        if (batch == null)
        {
            var activeSessionIds = SessionManager.ActiveSessionIds.ToList();
            var sessionCount = activeSessionIds.Count;
            var errorMessage = sessionCount switch
            {
                0 => $"Session '{sessionId}' not found. No active sessions exist. " +
                     "Possible causes: (1) Session was closed prematurely before completing operations, " +
                     "(2) Session never created. " +
                     "Recovery: Use excel_file(action='open') to create a new session.",
                1 => $"Session '{sessionId}' not found. Active session: {activeSessionIds[0]}. " +
                     "Did you close the session before completing all operations? Use the active sessionId shown above.",
                _ => $"Session '{sessionId}' not found. {sessionCount} active sessions exist. " +
                     "Verify you're using the correct sessionId from excel_file 'open' action."
            };
            throw new InvalidOperationException(errorMessage);
        }

        return await action(batch);
    }

    /// <summary>
    /// Throws exception for unknown actions.
    /// </summary>
    /// <param name="action">The invalid action that was attempted</param>
    /// <param name="supportedActions">List of supported actions for this tool</param>
    /// <exception cref="ArgumentException">Always throws with descriptive error message</exception>
    public static void ThrowUnknownAction(string action, params string[] supportedActions)
    {
        throw new ArgumentException(
            $"Unknown action '{action}'. Supported: {string.Join(", ", supportedActions)}", nameof(action));
    }

    /// <summary>
    /// Throws exception for missing required parameters.
    /// </summary>
    /// <param name="parameterName">Name of the missing parameter</param>
    /// <param name="action">The action that requires the parameter</param>
    /// <exception cref="ArgumentException">Always throws with descriptive error message</exception>
    public static void ThrowMissingParameter(string parameterName, string action)
    {
        throw new ArgumentException(
            $"{parameterName} is required for {action} action", parameterName);
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
    /// Executes a tool operation and serializes any exception using shared error formatting.
    /// Tracks tool usage telemetry (if enabled).
    /// </summary>
    /// <param name="toolName">Tool name for telemetry (e.g., "excel_range").</param>
    /// <param name="actionName">Action string (kebab-case) included in error context.</param>
    /// <param name="operation">Synchronous operation to execute.</param>
    /// <param name="customHandler">Optional handler that can override default error serialization. Return null/empty to fall back to default.</param>
    /// <returns>Serialized JSON response.</returns>
    public static string ExecuteToolAction(
        string toolName,
        string actionName,
        Func<string> operation,
        Func<Exception, string?>? customHandler = null) =>
        ExecuteToolAction(toolName, actionName, null, operation, customHandler);

    /// <summary>
    /// Executes a tool operation and serializes any exception using shared error formatting.
    /// Tracks tool usage telemetry (if enabled).
    /// </summary>
    /// <param name="toolName">Tool name for telemetry (e.g., "excel_range").</param>
    /// <param name="actionName">Action string (kebab-case) included in error context.</param>
    /// <param name="excelPath">Optional Excel path for context in error messages.</param>
    /// <param name="operation">Synchronous operation to execute.</param>
    /// <param name="customHandler">Optional handler that can override default error serialization. Return null/empty to fall back to default.</param>
    /// <returns>Serialized JSON response.</returns>
    public static string ExecuteToolAction(
        string toolName,
        string actionName,
        string? excelPath,
        Func<string> operation,
        Func<Exception, string?>? customHandler = null)
    {
        var stopwatch = Stopwatch.StartNew();
        var success = false;

        try
        {
            var result = operation();
            success = true;
            return result;
        }
        catch (Exception ex)
        {
            if (customHandler != null)
            {
                var custom = customHandler(ex);
                if (!string.IsNullOrWhiteSpace(custom))
                {
                    return custom!;
                }
            }

            return SerializeToolError(actionName, excelPath, ex);
        }
        finally
        {
            stopwatch.Stop();
            ExcelMcpTelemetry.TrackToolInvocation(toolName, actionName, stopwatch.ElapsedMilliseconds, success);
        }
    }

    /// <summary>
    /// Serializes a tool error response with consistent structure.
    /// </summary>
    /// <param name="actionName">Action string (kebab-case) included in message.</param>
    /// <param name="excelPath">Optional Excel path context.</param>
    /// <param name="ex">Exception to serialize.</param>
    /// <returns>Serialized JSON error payload.</returns>
    public static string SerializeToolError(string actionName, string? excelPath, Exception ex)
    {
        var errorMessage = excelPath != null
            ? $"{actionName} failed for '{excelPath}': {ex.Message}"
            : $"{actionName} failed: {ex.Message}";

        var payload = new
        {
            success = false,
            errorMessage,
            isError = true
        };

        return JsonSerializer.Serialize(payload, JsonOptions);
    }
}
