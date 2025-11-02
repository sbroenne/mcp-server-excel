using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP prompts for teaching LLMs about Power Query and Data Model workflows.
/// Critical upfront knowledge to prevent common mistakes.
/// </summary>
[McpServerPromptType]
public static class ExcelPowerQueryDataModelPrompts
{
    [McpServerPrompt(Name = "excel_powerquery_datamodel_guide")]
    [Description("Essential guide: Understanding Power Query load destinations and Data Model workflows")]
    public static ChatMessage PowerQueryDataModelGuide()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadPrompt("excel_powerquery.md"));
    }
}
