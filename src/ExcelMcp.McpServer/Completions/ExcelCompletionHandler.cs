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
        // Load destination completions for Power Query
        else if (argumentName == "loadDestination")
        {
            suggestions =
            [
                "worksheet", "data-model", "both", "connection-only"
            ];
        }
        // Action parameter completions for Parameter tool
        else if (argumentName == "action" && promptName.Contains("parameter", StringComparison.OrdinalIgnoreCase))
        {
            suggestions =
            [
                "list", "get", "set", "update", "create", "delete", "create-bulk"
            ];
        }
        // Action parameter completions for Data Model prompts
        else if (argumentName == "action" && promptName.Contains("datamodel", StringComparison.OrdinalIgnoreCase))
        {
            suggestions =
            [
                "list-tables", "list-measures", "view-measure", "export-measure",
                "list-relationships", "refresh", "delete-measure", "delete-relationship",
                "list-columns", "view-table", "get-model-info", "create-measure",
                "update-measure", "create-relationship", "update-relationship",
                "create-column", "view-column", "update-column", "delete-column", "validate-dax"
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
        // Action parameter completions for range prompts
        else if (argumentName == "action" && promptName.Contains("range", StringComparison.OrdinalIgnoreCase))
        {
            suggestions =
            [
                "get-values", "set-values", "get-formulas", "set-formulas",
                "clear-all", "clear-contents", "clear-formats",
                "copy", "copy-values", "copy-formulas",
                "insert-cells", "delete-cells", "insert-rows", "delete-rows",
                "insert-columns", "delete-columns",
                "find", "replace", "sort",
                "get-used-range", "get-current-region", "get-range-info",
                "add-hyperlink", "remove-hyperlink", "list-hyperlinks", "get-hyperlink",
                "format-range", "validate-range", "set-number-format", "get-number-formats"
            ];
        }
        // Action parameter completions for table prompts
        else if (argumentName == "action" && promptName.Contains("table", StringComparison.OrdinalIgnoreCase))
        {
            suggestions =
            [
                "list", "create", "resize", "rename", "delete", "add-column",
                "remove-column", "rename-column", "append-rows", "apply-filter",
                "clear-filters", "sort", "add-to-datamodel"
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
        // Format string completions for measures and ranges
        else if (argumentName == "formatString" || argumentName == "formatCode")
        {
            suggestions =
            [
                "#,##0.00",           // Standard number
                "$#,##0.00",          // Currency with decimals
                "$#,##0",             // Currency no decimals
                "0.00%",              // Percentage
                "#,##0",              // Whole number with thousands
                "mm/dd/yyyy",         // Short date
                "m/d/yyyy",           // Short date no leading zero
                "dddd, mmmm dd, yyyy", // Long date
                "h:mm AM/PM",         // Time
                "General Number",     // General
                "@"                   // Text
            ];
        }
        // Range address completions
        else if (argumentName == "rangeAddress")
        {
            suggestions =
            [
                "A1:Z100",     // Common data range
                "A1:E50",      // Medium table
                "A1:D10",      // Small table
                "A1",          // Single cell
                "A1:A1000",    // Single column
                "1:1",         // Entire first row
                "A:A",         // Entire column A
                "SalesData",   // Named range example
                "DataRange"    // Named range example
            ];
        }
        // Sheet name completions
        else if (argumentName == "sheetName" || argumentName == "targetSheet" || argumentName == "sourceSheet")
        {
            suggestions =
            [
                "Sheet1", "Data", "Report", "Summary", "Analysis",
                "Sales", "Products", "Customers", "Dashboard", "Settings"
            ];
        }
        // Validation type completions
        else if (argumentName == "validationType")
        {
            suggestions =
            [
                "list", "decimal", "whole", "date", "time", "textLength", "custom"
            ];
        }
        // Validation operator completions
        else if (argumentName == "validationOperator")
        {
            suggestions =
            [
                "between", "notBetween", "equal", "notEqual",
                "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual"
            ];
        }
        // Error style completions
        else if (argumentName == "errorStyle")
        {
            suggestions =
            [
                "stop", "warning", "information"
            ];
        }
        // Alignment completions
        else if (argumentName == "horizontalAlignment")
        {
            suggestions =
            [
                "left", "center", "right", "justify", "distributed"
            ];
        }
        else if (argumentName == "verticalAlignment")
        {
            suggestions =
            [
                "top", "center", "bottom", "justify", "distributed"
            ];
        }
        // Border style completions
        else if (argumentName == "borderStyle")
        {
            suggestions =
            [
                "none", "continuous", "dash", "dot", "double", "dashDot", "dashDotDot"
            ];
        }
        // Border weight completions
        else if (argumentName == "borderWeight")
        {
            suggestions =
            [
                "hairline", "thin", "medium", "thick"
            ];
        }
        // Color completions (common Excel theme colors)
        else if (argumentName == "fontColor" || argumentName == "fillColor" || argumentName == "borderColor")
        {
            suggestions =
            [
                "#000000",  // Black
                "#FFFFFF",  // White
                "#FF0000",  // Red
                "#00FF00",  // Green
                "#0000FF",  // Blue
                "#FFFF00",  // Yellow
                "#FFA500",  // Orange
                "#800080",  // Purple
                "#4472C4",  // Excel blue (headers)
                "#70AD47",  // Excel green
                "#ED7D31",  // Excel orange
                "#FFC000",  // Excel yellow
                "#5B9BD5",  // Light blue
                "#A5A5A5",  // Gray
                "#D3D3D3"   // Light gray
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

    private static List<string> GetResourceUriCompletions(string uri, string currentValue)
    {
        var suggestions = new List<string>();

        // Suggest Excel file paths for excel:// URIs or when completing file paths
        if (uri.StartsWith("excel://", StringComparison.OrdinalIgnoreCase) ||
            currentValue.Contains(".xlsx", StringComparison.OrdinalIgnoreCase) ||
            currentValue.Contains(".xlsm", StringComparison.OrdinalIgnoreCase) ||
            currentValue.Contains(":\\", StringComparison.OrdinalIgnoreCase))
        {
            // Scan common directories for Excel files
            var commonPaths = new[]
            {
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Desktop"),
                Environment.CurrentDirectory
            };

            foreach (var basePath in commonPaths)
            {
                try
                {
                    if (Directory.Exists(basePath))
                    {
                        var excelFiles = Directory.GetFiles(basePath, "*.xlsx", SearchOption.TopDirectoryOnly)
                            .Concat(Directory.GetFiles(basePath, "*.xlsm", SearchOption.TopDirectoryOnly))
                            .Take(5)
                            .ToList();

                        suggestions.AddRange(excelFiles);
                    }
                }
                catch
                {
                    // Ignore access denied or other errors
                }
            }
        }

        // Remove duplicates and limit to 15 suggestions
        suggestions = suggestions.Distinct().Take(15).ToList();

        // If no files found, provide example paths
        if (suggestions.Count == 0)
        {
            suggestions =
            [
                "C:\\Data\\workbook.xlsx",
                "C:\\Reports\\analysis.xlsx",
                "C:\\Projects\\dashboard.xlsx"
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
