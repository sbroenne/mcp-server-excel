namespace Sbroenne.ExcelMcp.Core.Attributes;

/// <summary>
/// Declares an additional case-insensitive external name for an enum member.
/// Generated CLI, service, and MCP parsing accepts the alias explicitly.
/// </summary>
[AttributeUsage(AttributeTargets.Field, AllowMultiple = true, Inherited = false)]
public sealed class EnumAliasAttribute : Attribute
{
    /// <summary>
    /// Creates an alias for an enum member.
    /// </summary>
    /// <param name="alias">External enum value accepted by generated parsers.</param>
    public EnumAliasAttribute(string alias)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(alias);
        Alias = alias;
    }

    /// <summary>
    /// External enum value accepted by generated parsers.
    /// </summary>
    public string Alias { get; }
}
