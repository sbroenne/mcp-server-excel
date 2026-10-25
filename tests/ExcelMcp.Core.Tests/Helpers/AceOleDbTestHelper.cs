using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Tests.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Helper functions for creating temporary Excel data sources consumable via the
/// Microsoft.ACE.OLEDB provider. These utilities are used by OLEDB integration tests.
/// </summary>
public static class AceOleDbTestHelper
{
    private const string ProviderName = "Microsoft.ACE.OLEDB.16.0";

    /// <summary>
    /// Creates a temporary Excel workbook containing a simple Products table that can
    /// be queried via ACE OLEDB. The workbook is saved to <paramref name="workbookPath"/>.
    /// </summary>
    public static void CreateExcelDataSource(string workbookPath)
    {
        SavedWorkbookTemplates.CopyBlankTo(workbookPath);
        using var batch = ExcelSession.BeginBatch(workbookPath);
        batch.Execute((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                sheet.Name = "Products";
                range = sheet.Range["A1:B3"];
                range.Value2 = new object[,]
                {
                    { "Product", "Price" }, { "Widget", 19.99 }, { "Gadget", 29.99 }
                };
                ctx.Book.Save();
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    /// <summary>
    /// Updates the value cells in the Excel data source using an Excel COM operation.
    /// Useful for simulating data source changes before calling Refresh on the connection.
    /// </summary>
    public static void UpdateExcelDataSource(string workbookPath, Action<Excel.Worksheet> updateAction)
    {
        using var batch = ExcelSession.BeginBatch(workbookPath);
        batch.Execute((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                updateAction(sheet);
                ctx.Book.Save();
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    /// <summary>
    /// Generates an ACE OLEDB connection string targeting the supplied Excel workbook.
    /// </summary>
    public static string GetExcelConnectionString(string workbookPath, bool headersPresent = true)
    {
        string hdr = headersPresent ? "YES" : "NO";
        return $"OLEDB;Provider={ProviderName};Data Source={workbookPath};Extended Properties=\"Excel 12.0 Xml;HDR={hdr}\"";
    }

    /// <summary>
    /// Returns the default SQL command text used in tests to select data from the Products worksheet.
    /// </summary>
    public static string GetDefaultCommandText() => "SELECT * FROM [Products$]";
}


