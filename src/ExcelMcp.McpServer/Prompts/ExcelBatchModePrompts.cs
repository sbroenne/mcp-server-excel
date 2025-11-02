using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// Essential guide for using Excel batch mode to achieve 95% faster operations.
/// </summary>
[McpServerPromptType]
public static class ExcelBatchModePrompts
{
    /// <summary>
    /// Essential guide: When and how to use Excel batch mode for 95% faster operations.
    /// </summary>
    [McpServerPrompt(Name = "excel_batch_mode_guide")]
    [Description("Essential guide: When and how to use Excel batch mode for 75-95% faster operations")]
    public static ChatMessage BatchModeGuide()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadPrompt("excel_batch.md"));
    }
}
