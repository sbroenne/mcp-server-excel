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
    /// Whether the action may change Excel or external state. Defaults to true so new actions fail closed.
    /// Read-only actions must explicitly set this to false.
    /// </summary>
    public bool IsMutation { get; set; } = true;

    /// <summary>
    /// Whether this action requires an existing Excel session. Defaults to true so new actions
    /// fail closed. Set to false for self-contained atomic actions that manage their own files.
    /// </summary>
    public bool RequiresSession { get; set; } = true;

    /// <summary>
    /// Creates a new ServiceActionAttribute.
    /// </summary>
    /// <param name="action">The action name in kebab-case (e.g., "get-load-config")</param>
    public ServiceActionAttribute(string action)
    {
        Action = action ?? throw new ArgumentNullException(nameof(action));
    }
}
