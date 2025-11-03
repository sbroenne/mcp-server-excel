using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP prompts for Excel range operations - values, formulas, formatting, validation.
/// </summary>
[McpServerPromptType]
public static class ExcelRangePrompts
{
    /// <summary>
    /// Guide for formatting and styling Excel ranges.
    /// </summary>
    [McpServerPrompt(Name = "excel_range_formatting_guide")]
    [Description("Best practices for formatting and styling Excel ranges")]
    public static ChatMessage FormattingGuide()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadPrompt("excel_range.md"));
    }

    /// <summary>
    /// Guide for data validation rules in Excel ranges.
    /// </summary>
    [McpServerPrompt(Name = "excel_range_validation_guide")]
    [Description("Guide for adding data validation rules to Excel ranges")]
    public static ChatMessage ValidationGuide()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadElicitation("data_validation.md"));
    }

    /// <summary>
    /// Complete workflow combining values, formulas, formatting, and validation.
    /// </summary>
    [McpServerPrompt(Name = "excel_range_complete_workflow")]
    [Description("Complete workflow: values → formulas → formatting → validation")]
    public static ChatMessage CompleteWorkflow()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadPrompt("excel_range.md"));
    }
}
