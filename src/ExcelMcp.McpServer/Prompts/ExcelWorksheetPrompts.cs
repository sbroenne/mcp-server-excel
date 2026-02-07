using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP prompts for teaching LLMs about Excel worksheet operations,
/// particularly cross-file copy and move operations.
/// </summary>
[McpServerPromptType]
public static class ExcelWorksheetPrompts
{
    [McpServerPrompt(Name = "excel_worksheet_cross_file_guide")]
    [Description("Guide for copying and moving worksheets between different Excel files using atomic operations")]
    public static ChatMessage CrossFileOperationsGuide()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadPrompt("excel_worksheet.md"));
    }
}


