using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Diagnostics.Tests;

internal static class DiagnosticExcelObjects
{
    internal static Excel.Workbook AddWorkbook(Excel.Application application)
    {
        Excel.Workbooks? books = null;
        try
        {
            books = application.Workbooks;
            return books.Add();
        }
        finally
        {
            ComUtilities.Release(ref books);
        }
    }

    internal static Excel.Workbook OpenWorkbook(Excel.Application application, string path)
    {
        Excel.Workbooks? books = null;
        try
        {
            books = application.Workbooks;
            return books.Open(path);
        }
        finally
        {
            ComUtilities.Release(ref books);
        }
    }

    internal static Excel.Worksheet GetWorksheet(Excel.Workbook book, object index)
    {
        Excel.Sheets? sheets = null;
        try
        {
            sheets = book.Worksheets;
            return (Excel.Worksheet)sheets[index];
        }
        finally
        {
            ComUtilities.Release(ref sheets);
        }
    }

    internal static Excel.Worksheet AddWorksheet(Excel.Workbook book)
    {
        Excel.Sheets? sheets = null;
        try
        {
            sheets = book.Worksheets;
            return (Excel.Worksheet)sheets.Add();
        }
        finally
        {
            ComUtilities.Release(ref sheets);
        }
    }

    internal static Excel.PivotTable GetPivotTable(Excel.Workbook book, string sheetName, string name)
    {
        Excel.Worksheet? sheet = null;
        try
        {
            sheet = GetWorksheet(book, sheetName);
            return (Excel.PivotTable)sheet.PivotTables(name);
        }
        finally
        {
            ComUtilities.Release(ref sheet);
        }
    }

    internal static Excel.PivotField GetDataField(Excel.PivotTable pivot, object index)
    {
        Excel.PivotFields? fields = null;
        try
        {
            fields = (Excel.PivotFields)pivot.DataFields;
            return fields.Item(index);
        }
        finally
        {
            ComUtilities.Release(ref fields);
        }
    }

    internal static object? ReadCell(Excel.Worksheet sheet, string address)
    {
        Excel.Range? range = null;
        try
        {
            range = sheet.Range[address];
            return (object?)range.Value2;
        }
        finally
        {
            ComUtilities.Release(ref range);
        }
    }
}
