using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP prompts for teaching LLMs about Excel range operations,
/// particularly number formatting with locale-awareness.
/// </summary>
[McpServerPromptType]
public static class ExcelRangePrompts
{
    [McpServerPrompt(Name = "range_number_formatting_guide")]
    [Description("Guide for number formatting: When to use SetNumberFormat (locale-aware) vs SetNumberFormatCustom (raw format codes)")]
    public static ChatMessage NumberFormattingGuide()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadPrompt("excel_range.md"));
    }
}


