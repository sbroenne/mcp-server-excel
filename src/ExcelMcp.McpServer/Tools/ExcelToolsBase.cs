using ModelContextProtocol.Server;
using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Diagnostics.CodeAnalysis;

#pragma warning disable IL2070 // 'this' argument does not satisfy 'DynamicallyAccessedMembersAttribute' requirements

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Base class for Excel MCP tools providing common patterns and utilities.
/// All Excel tools inherit from this to ensure consistency for LLM usage.
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
    /// Creates a standardized error response for unknown actions.
    /// Pattern: Use this for consistent error handling across all tools.
    /// </summary>
    /// <param name="action">The invalid action that was attempted</param>
    /// <param name="supportedActions">List of supported actions for this tool</param>
    /// <returns>JSON error response</returns>
    public static string CreateUnknownActionError(string action, params string[] supportedActions)
    {
        return JsonSerializer.Serialize(new 
        { 
            error = $"Unknown action '{action}'. Supported: {string.Join(", ", supportedActions)}" 
        }, JsonOptions);
    }

    /// <summary>
    /// Creates a standardized exception error response.
    /// Pattern: Use this for consistent exception handling across all tools.
    /// </summary>
    /// <param name="ex">The exception that occurred</param>
    /// <param name="action">The action that was being attempted</param>
    /// <param name="filePath">The file path involved (optional)</param>
    /// <returns>JSON error response</returns>
    public static string CreateExceptionError(Exception ex, string action, string? filePath = null)
    {
        var errorObj = new Dictionary<string, object?>
        {
            ["error"] = ex.Message,
            ["action"] = action
        };
        
        if (!string.IsNullOrEmpty(filePath))
        {
            errorObj["filePath"] = filePath;
        }

        return JsonSerializer.Serialize(errorObj, JsonOptions);
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