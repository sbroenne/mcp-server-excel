namespace Sbroenne.ExcelMcp.Core.Attributes;

/// <summary>Declares action applicability for a specialized MCP input.</summary>
[AttributeUsage(AttributeTargets.Parameter, AllowMultiple = true)]
public sealed class McpActionParameterAttribute(string action) : Attribute
{
    /// <summary>The action accepting this parameter.</summary>
    public string Action { get; } = action;

    /// <summary>Whether the action requires a non-null value.</summary>
    public bool Required { get; set; }
}
