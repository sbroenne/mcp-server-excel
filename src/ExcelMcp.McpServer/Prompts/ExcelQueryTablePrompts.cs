using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP prompts for Excel QueryTable management.
/// </summary>
[McpServerPromptType]
public static class ExcelQueryTablePrompts
{
    /// <summary>
    /// Guide for QueryTable operations and integration with connections and Power Query.
    /// </summary>
    [McpServerPrompt(Name = "excel_querytable_guide")]
    [Description("Guide for Excel QueryTable operations - simple data imports with reliable synchronous refresh")]
    public static ChatMessage QueryTableGuide()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadPrompt("excel_querytable.md"));
    }
}
