namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Indicates that retaining another workbook savepoint would exceed a configured limit.
/// </summary>
public sealed class WorkbookSavepointStorageLimitException : InvalidOperationException
{
    /// <summary>Initializes a savepoint storage limit failure.</summary>
    public WorkbookSavepointStorageLimitException(string message)
        : base(message)
    {
    }
}
