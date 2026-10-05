namespace Sbroenne.ExcelMcp.Core.Attributes;

/// <summary>
/// Exposes the named service actions through a separate read-only MCP tool.
/// The generated tool name is the interface MCP tool name with "_read" appended.
/// </summary>
[AttributeUsage(AttributeTargets.Interface, AllowMultiple = false, Inherited = false)]
public sealed class McpReadOnlyActionsAttribute(params string[] actionNames) : Attribute
{
    internal IReadOnlyList<string> ActionNames { get; } = actionNames;
}
