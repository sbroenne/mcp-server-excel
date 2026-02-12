using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP prompts for gathering required information before executing Excel operations.
/// Acts as pre-flight checklists to prevent back-and-forth with users.
/// </summary>
[McpServerPromptType]
public static class ExcelElicitationPrompts
{
    [McpServerPrompt(Name = "data_validation_checklist")]
    [Description("Information needed before adding data validation to ranges")]
    public static ChatMessage DataValidationChecklist()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadElicitation("data_validation.md"));
    }
}


