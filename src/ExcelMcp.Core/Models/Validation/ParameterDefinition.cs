using System.Text.RegularExpressions;

namespace Sbroenne.ExcelMcp.Core.Models.Validation;

/// <summary>
/// Defines validation rules for a parameter
/// </summary>
public class ParameterDefinition
{
    /// <summary>
    /// Parameter name
    /// </summary>
    public string Name { get; init; } = "";

    /// <summary>
    /// Whether parameter is required
    /// </summary>
    public bool Required { get; init; }

    /// <summary>
    /// Minimum length for string parameters
    /// </summary>
    public int? MinLength { get; init; }

    /// <summary>
    /// Maximum length for string parameters
    /// </summary>
    public int? MaxLength { get; init; }

    /// <summary>
    /// Regex pattern for string validation
    /// </summary>
    public string? Pattern { get; init; }

    /// <summary>
    /// Allowed file extensions (without dot, e.g., "xlsx", "xlsm")
    /// </summary>
    public string[]? FileExtensions { get; init; }

    /// <summary>
    /// Allowed values for enum-like parameters
    /// </summary>
    public string[]? AllowedValues { get; init; }

    /// <summary>
    /// Parameter description for help text
    /// </summary>
    public string? Description { get; init; }

    /// <summary>
    /// Default value if not provided
    /// </summary>
    public object? DefaultValue { get; init; }

    /// <summary>
    /// Validates a parameter value
    /// </summary>
    public ValidationResult Validate(object? value)
    {
        // Check required
        if (Required && value == null)
        {
            return ValidationResult.Failure(Name, $"Parameter '{Name}' is required");
        }

        if (value == null)
        {
            return ValidationResult.Success();
        }

        string? stringValue = value.ToString();

        // Check min length
        if (MinLength.HasValue && stringValue != null && stringValue.Length < MinLength.Value)
        {
            return ValidationResult.Failure(Name, 
                $"Parameter '{Name}' must be at least {MinLength.Value} characters");
        }

        // Check max length
        if (MaxLength.HasValue && stringValue != null && stringValue.Length > MaxLength.Value)
        {
            return ValidationResult.Failure(Name,
                $"Parameter '{Name}' must not exceed {MaxLength.Value} characters");
        }

        // Check pattern
        if (Pattern != null && stringValue != null)
        {
            if (!Regex.IsMatch(stringValue, Pattern))
            {
                return ValidationResult.Failure(Name,
                    $"Parameter '{Name}' has invalid format");
            }
        }

        // Check file extensions
        if (FileExtensions != null && stringValue != null)
        {
            string ext = Path.GetExtension(stringValue).TrimStart('.');
            if (!FileExtensions.Contains(ext, StringComparer.OrdinalIgnoreCase))
            {
                return ValidationResult.Failure(Name,
                    $"Parameter '{Name}' must have extension: {string.Join(", ", FileExtensions)}");
            }
        }

        // Check allowed values
        if (AllowedValues != null && stringValue != null)
        {
            if (!AllowedValues.Contains(stringValue, StringComparer.OrdinalIgnoreCase))
            {
                return ValidationResult.Failure(Name,
                    $"Parameter '{Name}' must be one of: {string.Join(", ", AllowedValues)}");
            }
        }

        return ValidationResult.Success();
    }
}
