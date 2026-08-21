namespace Sbroenne.ExcelMcp.Core.Attributes;

/// <summary>
/// Marks a required string parameter whose empty value has defined semantics.
/// The parameter must still be supplied by public callers.
/// </summary>
[AttributeUsage(AttributeTargets.Parameter)]
public sealed class AllowEmptyStringAttribute : Attribute
{
}
