using System.Text.Json.Nodes;

namespace Sbroenne.ExcelMcp.McpServer.Completions;

/// <summary>
/// Provides autocomplete suggestions for Excel MCP prompts and resources.
/// Implements completion support by handling the completion/complete JSON-RPC method
/// as described in https://devblogs.microsoft.com/dotnet/mcp-csharp-sdk-2025-06-18-update/
/// </summary>
public static class ExcelCompletionHandler
{
    /// <summary>
    /// Handle completion requests according to MCP spec.
    /// Returns suggestions for prompt arguments based on the argument name and context.
    /// </summary>
    public static JsonObject HandleCompletion(JsonObject request)
    {
        try
        {
            var paramsObj = request["params"] as JsonObject;
            if (paramsObj == null)
            {
                return CreateEmptyCompletion();
            }

            var refObj = paramsObj["ref"] as JsonObject;
            var argument = paramsObj["argument"] as JsonObject;

            if (refObj == null || argument == null)
            {
                return CreateEmptyCompletion();
            }

            var refType = refObj["type"]?.ToString();
            var argumentName = argument["name"]?.ToString();
            var argumentValue = argument["value"]?.ToString() ?? "";

            // Handle prompt argument completions
            if (refType == "ref/prompt")
            {
                var promptName = refObj["name"]?.ToString() ?? "";
                var suggestions = GetPromptArgumentCompletions(promptName, argumentName, argumentValue);
                return CreateCompletionResult(suggestions);
            }

            // Handle resource URI completions
            if (refType == "ref/resource")
            {
                var uri = refObj["uri"]?.ToString() ?? "";
                var suggestions = GetResourceUriCompletions(uri);
                return CreateCompletionResult(suggestions);
            }

            return CreateEmptyCompletion();
        }
        catch
        {
            return CreateEmptyCompletion();
        }
    }

    private static List<string> GetPromptArgumentCompletions(string promptName, string? argumentName, string currentValue)
    {
        var suggestions = new List<string>();

        // Action parameter completions for Power Query prompts
        if (argumentName == "action" && promptName.Contains("powerquery", StringComparison.OrdinalIgnoreCase))
        {
            suggestions =
            [
                "list", "view", "import", "export", "update", "delete", "refresh",
                "set-load-to-table", "set-load-to-data-model", "set-load-to-both",
                "set-connection-only", "get-load-config"
            ];
        }
        // Action parameter completions for VBA prompts
        else if (argumentName == "action" && promptName.Contains("vba", StringComparison.OrdinalIgnoreCase))
        {
            suggestions =
            [
                "list", "view", "export", "import", "update", "run", "delete"
            ];
        }
        // Action parameter completions for worksheet prompts
        else if (argumentName == "action" && promptName.Contains("worksheet", StringComparison.OrdinalIgnoreCase))
        {
            suggestions =
            [
                "list", "read", "write", "create", "rename", "copy", "delete", "clear", "append"
            ];
        }
        // Privacy level completions
        else if (argumentName == "privacyLevel")
        {
            suggestions =
            [
                "None", "Private", "Organizational", "Public"
            ];
        }

        // Filter suggestions based on current value (prefix matching)
        if (!string.IsNullOrEmpty(currentValue))
        {
            suggestions = suggestions
                .Where(s => s.StartsWith(currentValue, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        return suggestions;
    }

    private static List<string> GetResourceUriCompletions(string uri)
    {
        var suggestions = new List<string>();

        // Suggest Excel file paths for excel:// URIs
        if (uri.StartsWith("excel://", StringComparison.OrdinalIgnoreCase) ||
            uri.Contains(".xlsx", StringComparison.OrdinalIgnoreCase) ||
            uri.Contains(".xlsm", StringComparison.OrdinalIgnoreCase))
        {
            // Example suggestions - in a real implementation, could scan common directories
            suggestions =
            [
                "C:\\Data\\sales.xlsx",
                "C:\\Reports\\monthly-report.xlsx",
                "C:\\Analysis\\budget.xlsx"
            ];
        }

        return suggestions;
    }

    private static JsonObject CreateCompletionResult(List<string> suggestions)
    {
        var completion = new JsonObject
        {
            ["values"] = new JsonArray(suggestions.Select(s => JsonValue.Create(s)).ToArray()),
            ["total"] = suggestions.Count,
            ["hasMore"] = false
        };

        return new JsonObject
        {
            ["completion"] = completion
        };
    }

    private static JsonObject CreateEmptyCompletion()
    {
        return new JsonObject
        {
            ["completion"] = new JsonObject
            {
                ["values"] = new JsonArray(),
                ["total"] = 0,
                ["hasMore"] = false
            }
        };
    }
}
