namespace Sbroenne.ExcelMcp.Core.Attributes;

/// <summary>
/// Overrides the default action name derived from method name.
/// By default, action names are derived from method names using PascalCase → kebab-case convention.
/// Use this attribute only when the convention doesn't produce the desired action name.
/// </summary>
/// <remarks>
/// Convention: GetLoadConfig → "get-load-config"
/// Override example: [ServiceAction("custom-action")]
/// </remarks>
[AttributeUsage(AttributeTargets.Method, AllowMultiple = false, Inherited = false)]
public sealed class ServiceActionAttribute : Attribute
{
    /// <summary>
    /// The action name to use instead of the derived name.
    /// </summary>
    public string Action { get; }

    /// <summary>
    /// Creates a new ServiceActionAttribute.
    /// </summary>
    /// <param name="action">The action name in kebab-case (e.g., "get-load-config")</param>
    public ServiceActionAttribute(string action)
    {
        Action = action ?? throw new ArgumentNullException(nameof(action));
    }
}
