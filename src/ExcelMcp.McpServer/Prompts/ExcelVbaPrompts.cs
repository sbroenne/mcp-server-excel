using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP prompts for Excel VBA macro management.
/// </summary>
[McpServerPromptType]
public static class ExcelVbaPrompts
{
    /// <summary>
    /// Guide for VBA macro version control and automation workflows.
    /// </summary>
    [McpServerPrompt(Name = "excel_vba_version_control_guide")]
    [Description("Guide for managing VBA macros with version control and automation")]
    public static ChatMessage VbaVersionControlGuide()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadPrompt("excel_vba.md"));
    }
}
