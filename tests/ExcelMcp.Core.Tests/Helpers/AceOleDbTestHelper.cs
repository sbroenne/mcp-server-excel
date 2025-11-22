using Sbroenne.ExcelMcp.ComInterop.Session;

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
        ExcelSession.CreateNew(workbookPath, isMacroEnabled: false, (ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            sheet.Name = "Products";
            sheet.Range["A1"].Value2 = "Product";
            sheet.Range["B1"].Value2 = "Price";
            sheet.Range["A2"].Value2 = "Widget";
            sheet.Range["B2"].Value2 = 19.99;
            sheet.Range["A3"].Value2 = "Gadget";
            sheet.Range["B3"].Value2 = 29.99;

            ctx.Book.Save();
            return 0;
        });
    }

    /// <summary>
    /// Updates the value cells in the Excel data source using an Excel COM operation.
    /// Useful for simulating data source changes before calling Refresh on the connection.
    /// </summary>
    public static void UpdateExcelDataSource(string workbookPath, Action<dynamic> updateAction)
    {
        using var batch = ExcelSession.BeginBatch(workbookPath);
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            updateAction(sheet);
            ctx.Book.Save();
            return 0;
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
