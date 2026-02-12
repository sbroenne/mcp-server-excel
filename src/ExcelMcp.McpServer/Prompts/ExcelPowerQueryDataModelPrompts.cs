using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP prompts for Power Query M code syntax and Data Model DMV references.
/// Only content NOT already covered by tool descriptions.
/// </summary>
[McpServerPromptType]
public static class ExcelPowerQueryDataModelPrompts
{
    [McpServerPrompt(Name = "m_code_syntax")]
    [Description("Power Query M code syntax: column quoting rules, named range access, query chaining")]
    public static ChatMessage MCodeSyntax()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadPrompt("m_code_syntax.md"));
    }

    [McpServerPrompt(Name = "dmv_reference")]
    [Description("Data Model DMV query reference: working queries, limitations, and TMSCHEMA catalog for Excel's embedded Analysis Services")]
    public static ChatMessage DmvReference()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadPrompt("dmv_reference.md"));
    }
}


