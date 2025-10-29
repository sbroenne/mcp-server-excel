namespace Sbroenne.ExcelMcp.Core.Models.Validation;

/// <summary>
/// Result of parameter validation
/// </summary>
public class ValidationResult
{
    /// <summary>
    /// Whether validation passed
    /// </summary>
    public bool IsValid { get; init; }

    /// <summary>
    /// Error message if validation failed
    /// </summary>
    public string? ErrorMessage { get; init; }

    /// <summary>
    /// Name of parameter that failed validation
    /// </summary>
    public string? ParameterName { get; init; }

    /// <summary>
    /// Creates a successful validation result
    /// </summary>
    public static ValidationResult Success() => new() { IsValid = true };

    /// <summary>
    /// Creates a failed validation result
    /// </summary>
    public static ValidationResult Failure(string parameterName, string errorMessage) => new()
    {
        IsValid = false,
        ParameterName = parameterName,
        ErrorMessage = errorMessage
    };
}
