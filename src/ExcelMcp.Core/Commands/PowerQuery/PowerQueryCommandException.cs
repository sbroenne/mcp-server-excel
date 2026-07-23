namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Exception for Power Query refresh failures with a structured category.
/// </summary>
public sealed class PowerQueryCommandException : InvalidOperationException
{
    /// <summary>
    /// Initializes a new instance of the <see cref="PowerQueryCommandException"/> class.
    /// </summary>
    /// <param name="message">The Power Query error message.</param>
    /// <param name="errorCategory">The classified Power Query error category.</param>
    /// <param name="innerException">The underlying exception raised by Excel/COM.</param>
    public PowerQueryCommandException(string message, string errorCategory, Exception innerException)
        : base(message, innerException)
    {
        ErrorCategory = errorCategory;
    }

    /// <summary>
    /// Gets the classified Power Query error category.
    /// </summary>
    public string ErrorCategory { get; }
}
