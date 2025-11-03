using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP prompt for helping LLMs choose the right Excel tool for the task.
/// </summary>
[McpServerPromptType]
public static class ExcelToolSelectionPrompts
{
    /// <summary>
    /// Comprehensive guide for selecting the appropriate Excel tool.
    /// </summary>
    [McpServerPrompt(Name = "excel_tool_selection_guide")]
    [Description("Guide for choosing the right Excel tool based on the user's request")]
    public static ChatMessage ToolSelectionGuide()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadPrompt("tool_selection_guide.md"));
    }
}
