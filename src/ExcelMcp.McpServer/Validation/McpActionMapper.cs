using Sbroenne.ExcelMcp.Core.Models.Validation;

namespace Sbroenne.ExcelMcp.McpServer.Validation;

/// <summary>
/// Maps domain ActionDefinitions to MCP-specific metadata
/// This adapter lives in MCP layer, keeping validation layer client-agnostic
/// </summary>
public static class McpActionMapper
{
    /// <summary>
    /// MCP tool names by domain
    /// </summary>
    private static readonly Dictionary<string, string> McpToolNames = new()
    {
        ["PowerQuery"] = "excel_powerquery",
        ["Parameter"] = "excel_parameter",
        ["Table"] = "excel_table",
        ["DataModel"] = "excel_datamodel",
        ["VBA"] = "excel_vba",
        ["Connection"] = "excel_connection",
        ["Worksheet"] = "excel_worksheet",
        ["Range"] = "excel_range",
        ["File"] = "excel_file"
    };

    /// <summary>
    /// Gets MCP tool name for a domain
    /// </summary>
    public static string GetMcpToolName(string domain)
    {
        return McpToolNames.TryGetValue(domain, out var toolName) ? toolName : $"excel_{domain.ToLowerInvariant()}";
    }

    /// <summary>
    /// Gets MCP tool name for an action definition
    /// </summary>
    public static string GetMcpToolName(ActionDefinition action)
    {
        return GetMcpToolName(action.Domain);
    }

    /// <summary>
    /// Gets MCP action name (same as domain action name)
    /// </summary>
    public static string GetMcpActionName(ActionDefinition action)
    {
        return action.Name;
    }

    /// <summary>
    /// Generates regex pattern for all valid MCP actions in a domain
    /// </summary>
    public static string GetMcpActionRegex(IEnumerable<ActionDefinition> actions)
    {
        var actionNames = actions.Select(a => System.Text.RegularExpressions.Regex.Escape(a.Name));
        return $"^({string.Join("|", actionNames)})$";
    }
}
