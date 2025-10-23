using System.Diagnostics.CodeAnalysis;
using System.Text.Json;
using System.Text.Json.Serialization;
using ModelContextProtocol;
using Sbroenne.ExcelMcp.Core;

#pragma warning disable IL2070 // 'this' argument does not satisfy 'DynamicallyAccessedMembersAttribute' requirements

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Base class for Excel MCP tools providing common patterns and utilities.
/// All Excel tools inherit from this to ensure consistency for LLM usage.
/// Provides pooled Excel instance support for conversational workflow performance.
/// </summary>
[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicMethods)]
public static class ExcelToolsBase
{
    /// <summary>
    /// JSON serializer options with enum string conversion for user-friendly API responses.
    /// Used by all Excel tools for consistent JSON formatting.
    /// </summary>
    public static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() }
    };

    /// <summary>
    /// Executes an action with Excel, using pooled instance if available for better performance.
    /// This method provides automatic fallback to single-instance pattern if pooling is not enabled.
    ///
    /// Performance: Pooled instances reduce Excel startup overhead from ~2-5 seconds to near-instantaneous
    /// for cached workbooks in conversational MCP workflows.
    /// </summary>
    /// <typeparam name="T">Return type of the action</typeparam>
    /// <param name="filePath">Path to the Excel file</param>
    /// <param name="save">Whether to save changes to the file</param>
    /// <param name="action">Action to execute with Excel application and workbook</param>
    /// <returns>Result of the action</returns>
    public static T WithExcel<T>(string filePath, bool save, Func<dynamic, dynamic, T> action)
    {
        return ExcelToolsPoolManager.WithExcel(filePath, save, action);
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
    /// Special handling for ExcelPoolCapacityException to provide actionable guidance.
    /// </summary>
    /// <param name="ex">The exception that occurred</param>
    /// <param name="action">The action that was being attempted</param>
    /// <param name="filePath">The file path involved (optional)</param>
    /// <exception cref="McpException">Always throws with contextual error message</exception>
    public static void ThrowInternalError(Exception ex, string action, string? filePath = null)
    {
        // Special handling for pool capacity exception - provide LLM with actionable guidance
        if (ex is ExcelPoolCapacityException poolEx)
        {
            var message = $"{action} failed: Excel instance pool is at maximum capacity " +
                         $"({poolEx.ActiveInstances}/{poolEx.MaxInstances} instances active). " +
                         $"Idle instances are automatically cleaned up after {poolEx.IdleTimeout.TotalSeconds:F0} seconds. " +
                         $"\n\nSUGGESTED ACTIONS:\n" +
                         string.Join("\n", poolEx.SuggestedActions.Select((a, i) => $"{i + 1}. {a}"));

            throw new McpException(message, ex);
        }

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
