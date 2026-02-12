namespace Sbroenne.ExcelMcp.Core.Attributes;

/// <summary>
/// Specifies which MCP tool exposes this interface or method.
/// Used by code generator to group methods into MCP tools and generate MCP tool classes.
/// </summary>
/// <remarks>
/// Can be applied at interface level (all methods go to same tool)
/// or method level (methods can be split across different tools).
/// Method-level attribute overrides interface-level.
/// </remarks>
[AttributeUsage(AttributeTargets.Interface | AttributeTargets.Method, AllowMultiple = false, Inherited = false)]
public sealed class McpToolAttribute : Attribute
{
    /// <summary>
    /// The MCP tool name (e.g., "powerquery", "range").
    /// </summary>
    public string ToolName { get; }

    /// <summary>
    /// Human-readable title for the MCP tool (e.g., "Power Query Operations").
    /// Used in [McpServerTool(Title = ...)].
    /// </summary>
    public string? Title { get; set; }

    /// <summary>
    /// Whether the tool is destructive (modifies data). Default: true.
    /// Used in [McpServerTool(Destructive = ...)].
    /// </summary>
    public bool Destructive { get; set; } = true;

    /// <summary>
    /// MCP meta category for the tool (e.g., "data", "analysis", "query", "settings").
    /// Used in [McpMeta("category", ...)].
    /// </summary>
    public string? Category { get; set; }

    /// <summary>
    /// Tool description shown to LLMs via [Description("...")].
    /// Since source generators can't read XML docs from metadata references,
    /// this provides the description that appears in the MCP JSON schema.
    /// </summary>
    public string? Description { get; set; }

    /// <summary>
    /// Creates a new McpToolAttribute.
    /// </summary>
    /// <param name="toolName">The MCP tool name (e.g., "powerquery")</param>
    public McpToolAttribute(string toolName)
    {
        ToolName = toolName ?? throw new ArgumentNullException(nameof(toolName));
    }
}
