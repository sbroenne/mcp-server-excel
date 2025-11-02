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
    [McpServerPrompt(Name = "excel_powerquery_checklist")]
    [Description("Checklist of information needed before importing a Power Query")]
    public static ChatMessage PowerQueryChecklist()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadElicitation("powerquery_import.md"));
    }

    [McpServerPrompt(Name = "excel_dax_measure_checklist")]
    [Description("Information needed before creating DAX measures")]
    public static ChatMessage DaxMeasureChecklist()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadElicitation("dax_measure.md"));
    }

    [McpServerPrompt(Name = "excel_range_formatting_checklist")]
    [Description("Information needed before formatting Excel ranges")]
    public static ChatMessage RangeFormattingChecklist()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadElicitation("range_formatting.md"));
    }

    [McpServerPrompt(Name = "excel_data_validation_checklist")]
    [Description("Information needed before adding data validation to ranges")]
    public static ChatMessage DataValidationChecklist()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadElicitation("data_validation.md"));
    }

    [McpServerPrompt(Name = "excel_batch_mode_detection")]
    [Description("Guide for detecting when to use batch mode based on user request keywords")]
    public static ChatMessage BatchModeDetection()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadElicitation("batch_workflow.md"));
    }

    [McpServerPrompt(Name = "excel_troubleshooting_guide")]
    [Description("Common Excel MCP server issues and solutions")]
    public static ChatMessage TroubleshootingGuide()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadPrompt("server_quirks.md"));
    }
}
