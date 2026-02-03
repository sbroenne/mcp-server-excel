using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP prompts for teaching LLMs about Excel conditional formatting and slicer operations.
/// </summary>
[McpServerPromptType]
public static class ExcelConditionalFormatSlicerPrompts
{
    [McpServerPrompt(Name = "excel_conditionalformat_guide")]
    [Description("Conditional formatting: rule types, color scales, icon sets, and data bars")]
    public static ChatMessage ConditionalFormatGuide()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadPrompt("excel_conditionalformat.md"));
    }

    [McpServerPrompt(Name = "excel_slicer_guide")]
    [Description("Slicers: filtering PivotTables and Excel Tables with visual controls")]
    public static ChatMessage SlicerGuide()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadPrompt("excel_slicer.md"));
    }
}
