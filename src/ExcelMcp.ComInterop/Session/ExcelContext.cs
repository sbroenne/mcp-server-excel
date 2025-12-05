using Sbroenne.ExcelMcp.ComInterop.Formatting;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

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

        // Initialize number format translator with locale-specific codes from Excel
        FormatTranslator = new NumberFormatTranslator(excel);
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

    /// <summary>
    /// Gets the number format translator for converting US format codes to locale-specific codes.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Use this to translate format strings like "m/d/yyyy" or "$#,##0.00" to locale-specific codes
    /// (e.g., "M/T/JJJJ" and "$#.##0,00" on German Excel) before setting <c>Range.NumberFormat</c>.
    /// </para>
    /// <example>
    /// <code>
    /// string localeFormat = ctx.FormatTranslator.TranslateToLocale("m/d/yyyy");
    /// range.NumberFormat = localeFormat;
    /// </code>
    /// </example>
    /// </remarks>
    public NumberFormatTranslator FormatTranslator { get; }
}
