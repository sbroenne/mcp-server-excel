namespace Sbroenne.ExcelMcp.Core.Attributes;

/// <summary>
/// Indicates that this parameter is required for the action.
/// The generator will emit validation code that throws if the value is null/empty.
/// </summary>
[AttributeUsage(AttributeTargets.Parameter, AllowMultiple = false, Inherited = false)]
public sealed class RequiredParameterAttribute : Attribute
{
    /// <summary>
    /// Creates a RequiredParameter attribute.
    /// </summary>
    public RequiredParameterAttribute() { }
}
