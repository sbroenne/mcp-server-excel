using System.Diagnostics.CodeAnalysis;
using System.Text.Json;
using System.Text.Json.Serialization;
using ModelContextProtocol;
using Sbroenne.ExcelMcp.Core;
using Sbroenne.ExcelMcp.Core.Session;

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
    /// Executes an async Core command with batch session management.
    /// If batchId is provided, uses existing batch session. Otherwise, creates batch-of-one.
    /// 
    /// This is the standard pattern for all MCP tools to support both:
    /// - LLM-controlled batch sessions (pass batchId for multi-operation workflows)
    /// - Single operations (no batchId = automatic batch-of-one for backward compat)
    /// </summary>
    /// <typeparam name="T">Return type of the command</typeparam>
    /// <param name="batchId">Optional batch session ID from begin_excel_batch</param>
    /// <param name="filePath">Path to the Excel file (required if no batchId)</param>
    /// <param name="save">Whether to save changes (only used for batch-of-one)</param>
    /// <param name="action">Async action that takes IExcelBatch and returns Task&lt;T&gt;</param>
    /// <returns>Result of the command</returns>
    public static async Task<T> WithBatchAsync<T>(
        string? batchId,
        string filePath,
        bool save,
        Func<IExcelBatch, Task<T>> action)
    {
        if (!string.IsNullOrEmpty(batchId))
        {
            // Use existing batch session (LLM-controlled lifecycle)
            var batch = BatchSessionTool.GetBatch(batchId);
            if (batch == null)
            {
                throw new ModelContextProtocol.McpException(
                    $"Batch session '{batchId}' not found. It may have already been committed or never existed.");
            }
            
            // Verify file path matches batch
            if (!string.Equals(batch.WorkbookPath, Path.GetFullPath(filePath), StringComparison.OrdinalIgnoreCase))
            {
                throw new ModelContextProtocol.McpException(
                    $"File path mismatch. Batch session is for '{batch.WorkbookPath}' but operation requested '{filePath}'.");
            }
            
            return await action(batch);
        }
        else
        {
            // Batch-of-one (backward compatibility for single operations)
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await action(batch);
            
            if (save)
            {
                await batch.SaveAsync();
            }
            
            return result;
        }
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
