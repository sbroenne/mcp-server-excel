namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// A known operation failure with a stable category and its original context.
/// </summary>
public sealed class OperationFailureException : InvalidOperationException
{
    /// <summary>Creates a categorized failure without discarding its underlying cause.</summary>
    public OperationFailureException(
        OperationFailureCategory errorCategory, string? message, Exception? innerException = null)
        : base(message, innerException)
    {
        if (!Enum.IsDefined(errorCategory))
        {
            throw new ArgumentOutOfRangeException(nameof(errorCategory), errorCategory, "Unknown failure category.");
        }

        ErrorCategory = errorCategory;
    }

    /// <summary>Gets the known condition that prevented the operation.</summary>
    public OperationFailureCategory ErrorCategory { get; }
}
