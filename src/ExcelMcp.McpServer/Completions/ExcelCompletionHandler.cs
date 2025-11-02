using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.McpServer.Prompts;

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
                var suggestions = GetResourceUriCompletions(uri, argumentValue);
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

        // Action parameter completions - load from markdown files
        if (argumentName == "action")
        {
            if (promptName.Contains("powerquery", StringComparison.OrdinalIgnoreCase))
            {
                suggestions = MarkdownLoader.LoadCompletionValues("action_powerquery.md");
            }
            else if (promptName.Contains("parameter", StringComparison.OrdinalIgnoreCase) || 
                     promptName.Contains("namedrange", StringComparison.OrdinalIgnoreCase))
            {
                suggestions = MarkdownLoader.LoadCompletionValues("action_parameter.md");
            }
            else if (promptName.Contains("datamodel", StringComparison.OrdinalIgnoreCase))
            {
                suggestions = MarkdownLoader.LoadCompletionValues("action_datamodel.md");
            }
            else if (promptName.Contains("vba", StringComparison.OrdinalIgnoreCase))
            {
                suggestions = MarkdownLoader.LoadCompletionValues("action_vba.md");
            }
            else if (promptName.Contains("worksheet", StringComparison.OrdinalIgnoreCase))
            {
                suggestions = MarkdownLoader.LoadCompletionValues("action_worksheet.md");
            }
            else if (promptName.Contains("range", StringComparison.OrdinalIgnoreCase))
            {
                suggestions = MarkdownLoader.LoadCompletionValues("action_range.md");
            }
            else if (promptName.Contains("table", StringComparison.OrdinalIgnoreCase))
            {
                suggestions = MarkdownLoader.LoadCompletionValues("action_table.md");
            }
            else if (promptName.Contains("connection", StringComparison.OrdinalIgnoreCase))
            {
                suggestions = MarkdownLoader.LoadCompletionValues("action_connection.md");
            }
            else if (promptName.Contains("pivottable", StringComparison.OrdinalIgnoreCase))
            {
                suggestions = MarkdownLoader.LoadCompletionValues("action_pivottable.md");
            }
            else if (promptName.Contains("batch", StringComparison.OrdinalIgnoreCase))
            {
                suggestions = MarkdownLoader.LoadCompletionValues("action_batch.md");
            }
        }
        // Load destination completions for Power Query
        else if (argumentName == "loadDestination")
        {
            suggestions = MarkdownLoader.LoadCompletionValues("load_destination.md");
        }
        // Privacy level completions
        else if (argumentName == "privacyLevel")
        {
            suggestions = MarkdownLoader.LoadCompletionValues("privacy_level.md");
        }
        // Format code completions
        else if (argumentName == "formatCode" || argumentName == "formatString" || argumentName == "numberFormat")
        {
            suggestions = MarkdownLoader.LoadCompletionValues("format_codes.md");
        }
        // Validation type completions
        else if (argumentName == "validationType")
        {
            suggestions = MarkdownLoader.LoadCompletionValues("validation_types.md");
        }
        // Validation operator completions
        else if (argumentName == "validationOperator")
        {
            suggestions = MarkdownLoader.LoadCompletionValues("validation_operators.md");
        }
        // Error style completions
        else if (argumentName == "errorStyle")
        {
            suggestions = MarkdownLoader.LoadCompletionValues("error_styles.md");
        }
        // Alignment completions
        else if (argumentName == "horizontalAlignment")
        {
            suggestions = MarkdownLoader.LoadCompletionValues("alignment_horizontal.md");
        }
        else if (argumentName == "verticalAlignment")
        {
            suggestions = MarkdownLoader.LoadCompletionValues("alignment_vertical.md");
        }
        // Border style completions
        else if (argumentName == "borderStyle")
        {
            suggestions = MarkdownLoader.LoadCompletionValues("border_styles.md");
        }
        // Border weight completions
        else if (argumentName == "borderWeight")
        {
            suggestions = MarkdownLoader.LoadCompletionValues("border_weights.md");
        }
        // Color completions (common Excel theme colors)
        else if (argumentName == "fontColor" || argumentName == "fillColor" || argumentName == "borderColor")
        {
            suggestions = MarkdownLoader.LoadCompletionValues("colors_common.md");
        }

        return FilterSuggestions(suggestions, currentValue);
    }

    private static List<string> GetResourceUriCompletions(string uri, string currentValue)
    {
        var suggestions = new List<string>();

        // Excel file path completions
        if (uri.StartsWith("file://", StringComparison.OrdinalIgnoreCase) || 
            uri.Contains(".xlsx", StringComparison.OrdinalIgnoreCase) ||
            uri.Contains(".xlsm", StringComparison.OrdinalIgnoreCase))
        {
            suggestions =
            [
                "C:\\Data\\workbook.xlsx",
                "C:\\Reports\\financial-report.xlsx",
                "C:\\Projects\\analysis.xlsm",
                "workbook.xlsx",
                "report.xlsx"
            ];
        }

        return FilterSuggestions(suggestions, currentValue);
    }

    private static List<string> FilterSuggestions(List<string> suggestions, string currentValue)
    {
        if (string.IsNullOrWhiteSpace(currentValue))
        {
            return suggestions;
        }

        // Case-insensitive prefix matching
        return suggestions
            .Where(s => s.StartsWith(currentValue, StringComparison.OrdinalIgnoreCase))
            .ToList();
    }

    private static JsonObject CreateCompletionResult(List<string> suggestions)
    {
        var values = new JsonArray();
        foreach (var suggestion in suggestions)
        {
            var item = new JsonObject
            {
                ["value"] = suggestion,
                ["description"] = $"Autocomplete: {suggestion}"
            };
            values.Add(item);
        }

        return new JsonObject
        {
            ["values"] = values,
            ["total"] = suggestions.Count,
            ["hasMore"] = false
        };
    }

    private static JsonObject CreateEmptyCompletion()
    {
        return new JsonObject
        {
            ["values"] = new JsonArray(),
            ["total"] = 0,
            ["hasMore"] = false
        };
    }
}
