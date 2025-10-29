using System.Text.RegularExpressions;

namespace Sbroenne.ExcelMcp.Core.Models.Validation;

/// <summary>
/// Defines an action with its parameters and validation rules
/// Domain-focused definition without client-specific concerns
/// </summary>
public class ActionDefinition
{
    /// <summary>
    /// Domain/category of action (e.g., "PowerQuery", "Parameter", "Table")
    /// </summary>
    public string Domain { get; init; } = "";

    /// <summary>
    /// Action name (e.g., "list", "view", "create")
    /// </summary>
    public string Name { get; init; } = "";

    /// <summary>
    /// Parameter definitions for this action
    /// </summary>
    public ParameterDefinition[] Parameters { get; init; } = Array.Empty<ParameterDefinition>();

    /// <summary>
    /// Description of what this action does
    /// </summary>
    public string? Description { get; init; }

    /// <summary>
    /// Validates parameters for this action
    /// </summary>
    public ValidationResult ValidateParameters(Dictionary<string, object?> parameters)
    {
        foreach (var paramDef in Parameters)
        {
            parameters.TryGetValue(paramDef.Name, out var value);
            var result = paramDef.Validate(value);
            if (!result.IsValid)
            {
                return result;
            }
        }

        return ValidationResult.Success();
    }

    /// <summary>
    /// Gets parameter names for documentation/help
    /// </summary>
    public string[] GetRequiredParameterNames()
    {
        return Parameters.Where(p => p.Required).Select(p => p.Name).ToArray();
    }

    /// <summary>
    /// Gets parameter names for documentation/help
    /// </summary>
    public string[] GetOptionalParameterNames()
    {
        return Parameters.Where(p => !p.Required).Select(p => p.Name).ToArray();
    }
}
