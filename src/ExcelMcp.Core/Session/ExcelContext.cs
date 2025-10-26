namespace Sbroenne.ExcelMcp.Core.Session;

/// <summary>
/// Provides access to Excel COM objects for operations.
/// Simplifies passing Excel application and workbook to operations.
/// </summary>
public sealed class ExcelContext
{
    /// <summary>
    /// Creates a new ExcelContext.
    /// </summary>
    /// <param name="workbookPath">Full path to the workbook</param>
    /// <param name="excel">Excel.Application COM object</param>
    /// <param name="workbook">Excel.Workbook COM object</param>
    public ExcelContext(string workbookPath, dynamic excel, dynamic workbook)
    {
        WorkbookPath = workbookPath ?? throw new ArgumentNullException(nameof(workbookPath));
        App = excel ?? throw new ArgumentNullException(nameof(excel));
        Book = workbook ?? throw new ArgumentNullException(nameof(workbook));
    }

    /// <summary>
    /// Gets the full path to the workbook.
    /// </summary>
    public string WorkbookPath { get; }

    /// <summary>
    /// Gets the Excel.Application COM object.
    /// </summary>
    public dynamic App { get; }

    /// <summary>
    /// Gets the Excel.Workbook COM object.
    /// </summary>
    public dynamic Book { get; }
}
