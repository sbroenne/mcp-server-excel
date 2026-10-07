namespace Sbroenne.ExcelMcp.Core.Attributes;

/// <summary>
/// Marks an interface as a service category for code generation.
/// The command group name used by MCP routing, CLI commands, and batch files is derived
/// from the <see cref="McpToolAttribute"/> name with underscores removed
/// (e.g., "calculation_mode" → "calculationmode"). Interfaces without an MCP tool use
/// the lowercased <see cref="PascalName"/> (e.g., "Sheet" → "sheet").
/// </summary>
[AttributeUsage(AttributeTargets.Interface, AllowMultiple = false, Inherited = false)]
public sealed class ServiceCategoryAttribute : Attribute
{
    /// <summary>
    /// PascalCase name used for generated types (e.g., "PowerQuery" → PowerQueryAction).
    /// </summary>
    public string PascalName { get; }

    /// <summary>
    /// Creates a new ServiceCategoryAttribute.
    /// </summary>
    /// <param name="pascalName">PascalCase name used for generated types (e.g., "PowerQuery")</param>
    public ServiceCategoryAttribute(string pascalName)
    {
        PascalName = pascalName ?? throw new ArgumentNullException(nameof(pascalName));
    }
}
