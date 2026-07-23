namespace Sbroenne.ExcelMcp.Core.Attributes;

/// <summary>
/// Indicates that this enum parameter should be exposed as a string in MCP/CLI.
/// The generator will create a string parameter and generate parsing code.
/// Example: [FromString] PowerQueryLoadMode loadMode → generates string loadDestination parameter
/// </summary>
[AttributeUsage(AttributeTargets.Parameter, AllowMultiple = false, Inherited = false)]
public sealed class FromStringAttribute : Attribute
{
    /// <summary>
    /// Optional custom parameter name for the exposed string.
    /// If not specified, uses the original parameter name.
    /// Example: loadMode with ExposedName="loadDestination" → string loadDestination
    /// </summary>
    public string? ExposedName { get; }

    /// <summary>
    /// Creates a FromString attribute using the original parameter name.
    /// </summary>
    public FromStringAttribute() : this(null) { }

    /// <summary>
    /// Creates a FromString attribute with a custom exposed name.
    /// </summary>
    /// <param name="exposedName">The parameter name to expose in MCP/CLI</param>
    public FromStringAttribute(string? exposedName)
    {
        ExposedName = exposedName;
    }
}
