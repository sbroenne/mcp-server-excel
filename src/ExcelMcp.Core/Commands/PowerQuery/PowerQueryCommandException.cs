using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Exception for Power Query failures with a category and optional retained changes.
/// </summary>
public sealed class PowerQueryCommandException : InvalidOperationException
{
    /// <summary>
    /// Initializes a new instance of the <see cref="PowerQueryCommandException"/> class.
    /// </summary>
    /// <param name="message">The Power Query error message.</param>
    /// <param name="errorCategory">The classified Power Query error category.</param>
    /// <param name="innerException">The underlying exception raised by Excel/COM.</param>
    /// <param name="partialState">Changes retained after a failed create or update.</param>
    public PowerQueryCommandException(
        string message,
        string? errorCategory,
        Exception innerException,
        PowerQueryPartialState? partialState = null)
        : base(message, innerException)
    {
        ErrorCategory = errorCategory;
        PartialState = partialState;
    }

    /// <summary>
    /// Gets the classified Power Query error category.
    /// </summary>
    public string? ErrorCategory { get; }

    /// <summary>Gets the changes retained after failure, when applicable.</summary>
    public PowerQueryPartialState? PartialState { get; }
}
