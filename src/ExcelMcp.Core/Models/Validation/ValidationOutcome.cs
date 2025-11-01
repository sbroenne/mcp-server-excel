namespace Sbroenne.ExcelMcp.Core.Models.Validation;

/// <summary>
/// Represents the outcome of a validation operation
/// </summary>
public class ValidationOutcome
{
    public bool IsValid { get; init; }
    public string? ErrorMessage { get; init; }
    
    private ValidationOutcome(bool isValid, string? errorMessage = null)
    {
        IsValid = isValid;
        ErrorMessage = errorMessage;
    }
    
    public static ValidationOutcome Success() => new(true);
    
    public static ValidationOutcome Failure(string errorMessage) => new(false, errorMessage);
}
