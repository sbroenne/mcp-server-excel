using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP prompts for Excel parameter (named range) management.
/// </summary>
[McpServerPromptType]
public static class ExcelNamedRangePrompts
{
    /// <summary>
    /// Guide for efficient parameter creation using bulk operations.
    /// </summary>
    [McpServerPrompt(Name = "excel_namedrange_bulk_guide")]
    [Description("Guide for creating multiple Excel parameters efficiently using bulk operations")]
    public static ChatMessage NamedRangeBulkGuide()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadPrompt("excel_namedrange.md"));
    }
}
