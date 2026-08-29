namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Failure from atomic range-to-table conversion with structured rollback details.
/// </summary>
public sealed class TableRangeConversionException : InvalidOperationException
{
    /// <summary>
    /// Creates a conversion exception.
    /// </summary>
    public TableRangeConversionException(
        string message,
        TableRangeConversionFailureDetails details,
        Exception? innerException = null)
        : base(message, innerException)
    {
        Details = details;
    }

    /// <summary>
    /// Structured conversion and rollback details.
    /// </summary>
    public TableRangeConversionFailureDetails Details { get; }
}
