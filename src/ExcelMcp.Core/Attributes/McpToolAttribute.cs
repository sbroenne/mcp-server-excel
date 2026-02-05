namespace Sbroenne.ExcelMcp.Core.Attributes;

/// <summary>
/// Specifies which MCP tool exposes this interface or method.
/// Used by code generator to group methods into MCP tools.
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
    /// The MCP tool name (e.g., "excel_powerquery", "excel_range").
    /// </summary>
    public string ToolName { get; }

    /// <summary>
    /// Creates a new McpToolAttribute.
    /// </summary>
    /// <param name="toolName">The MCP tool name (e.g., "excel_powerquery")</param>
    public McpToolAttribute(string toolName)
    {
        ToolName = toolName ?? throw new ArgumentNullException(nameof(toolName));
    }
}
