namespace Sbroenne.ExcelMcp.Core.Models.Validation;

/// <summary>
/// Defines validation rules for an action parameter
/// </summary>
public class ActionParameter
{
    public string Name { get; init; } = string.Empty;
    public bool Required { get; init; }
    public string? Type { get; init; }
    public int? MinLength { get; init; }
    public int? MaxLength { get; init; }
    public string? Pattern { get; init; }
    public string[]? AllowedValues { get; init; }
    public string? Description { get; init; }
    
    public ValidationOutcome Validate(object? value)
    {
        if (Required && value == null)
            return ValidationOutcome.Failure($"Parameter '{Name}' is required");
        
        if (value == null)
            return ValidationOutcome.Success();
        
        var stringValue = value.ToString() ?? string.Empty;
        
        if (MinLength.HasValue && stringValue.Length < MinLength.Value)
            return ValidationOutcome.Failure($"Parameter '{Name}' must be at least {MinLength.Value} characters");
        
        if (MaxLength.HasValue && stringValue.Length > MaxLength.Value)
            return ValidationOutcome.Failure($"Parameter '{Name}' must not exceed {MaxLength.Value} characters");
        
        if (AllowedValues != null && AllowedValues.Length > 0)
        {
            if (!AllowedValues.Contains(stringValue))
                return ValidationOutcome.Failure($"Parameter '{Name}' must be one of: {string.Join(", ", AllowedValues)}");
        }
        
        return ValidationOutcome.Success();
    }
}
