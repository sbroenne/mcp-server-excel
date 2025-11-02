using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// Quick reference prompt for Excel connection types and COM API limitations.
/// </summary>
[McpServerPromptType]
public static class ExcelConnectionPrompts
{
    /// <summary>
    /// Quick reference for Excel connection types and critical COM API limitations.
    /// </summary>
    [McpServerPrompt(Name = "excel_connection_reference")]
    [Description("Quick reference: Excel connection types, which ones work via COM API, and critical limitations")]
    public static ChatMessage ConnectionReference()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadPrompt("excel_connection.md"));
    }
}
