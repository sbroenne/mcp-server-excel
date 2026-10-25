using System.Text.RegularExpressions;
using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

public partial class PivotTableCommands
{
    private static void ValidateCacheSharingInputs(string? shareCacheWith, string? shareCacheWithSheet)
    {
        if (shareCacheWith is null && shareCacheWithSheet is null)
            return;
        if (string.IsNullOrWhiteSpace(shareCacheWith) || string.IsNullOrWhiteSpace(shareCacheWithSheet))
            throw new ArgumentException("Specify both shareCacheWith and shareCacheWithSheet as nonblank strings, or omit both.");
    }

    private static Excel.PivotCache ResolveSharedCreationCache(
        Excel.Workbook book, Excel.Application app, string sheetName, string pivotName,
        Excel.Worksheet sourceSheet, Excel.Range sourceRange, Excel.ListObject? sourceTable)
    {
        Excel.PivotTable? pivot = null;
        Excel.PivotCache? cache = null;
        try
        {
            pivot = FindPivotTable(book, sheetName, pivotName);
            cache = pivot.PivotCache();
            if (cache.OLAP || cache.SourceType != Excel.XlPivotTableSourceType.xlDatabase)
                throw new NotSupportedException(
                    "Cache sharing supports worksheet-backed regular PivotTables only, not external or OLAP/Data Model caches. No PivotTable or cache was created.");
            if (cache.SourceData is not string reference ||
                !MatchesCreationSource(book, app, reference, sourceSheet, sourceRange, sourceTable))
                throw new ArgumentException(
                    "Cannot share the cache: the requested source does not match the existing worksheet-backed source identity. Table and fixed-range sources are distinct even when their cells match. No PivotTable or cache was created.");

            var result = cache;
            cache = null;
            return result;
        }
        finally
        {
            ComUtilities.Release(ref cache);
            ComUtilities.Release(ref pivot);
        }
    }

    internal static bool MatchesCreationSource(Excel.Workbook book, Excel.Application app, string reference,
        Excel.Worksheet sourceSheet, Excel.Range sourceRange, Excel.ListObject? sourceTable)
    {
        int separator = reference.LastIndexOf('!');
        string address = separator < 0 ? reference : reference[(separator + 1)..];
        if (separator >= 0 && !MatchesSourceWorksheet(book.Name, book.FullName, reference[..separator], sourceSheet.Name))
            return false;

        // Excel normalizes sheet!table to the table name, retaining the table's expanding source semantics.
        if (sourceTable is not null)
            return string.Equals(address, sourceTable.Name, StringComparison.OrdinalIgnoreCase);
        if (separator < 0)
            return false;

        // Native whole-axis references such as C1:C4 and R1:R6 also look like A1 cells.
        if (IsNativeR1C1SourceAddress(address))
            address = Convert.ToString((object)app.ConvertFormula(address, Excel.XlReferenceStyle.xlR1C1,
                Excel.XlReferenceStyle.xlA1, Excel.XlReferenceType.xlAbsolute),
                System.Globalization.CultureInfo.InvariantCulture)!;
        else if (!IsFixedA1SourceAddress(address))
            return false;

        Excel.Range? existingSource = null;
        try
        {
            existingSource = sourceSheet.Range[address];
            return string.Equals(
                existingSource.Address[true, true, Excel.XlReferenceStyle.xlA1, true],
                sourceRange.Address[true, true, Excel.XlReferenceStyle.xlA1, true],
                StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            ComUtilities.Release(ref existingSource);
        }
    }

    internal static bool IsNativeR1C1SourceAddress(string address) =>
        Regex.IsMatch(address,
            @"^(R[0-9]+C[0-9]+(:R[0-9]+C[0-9]+)?|C[0-9]+(:C[0-9]+)?|R[0-9]+(:R[0-9]+)?)$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));

    private static bool IsFixedA1SourceAddress(string address) =>
        Regex.IsMatch(address,
            @"^(\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?|\$?[A-Z]+:\$?[A-Z]+|\$?[0-9]+:\$?[0-9]+)$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));

    internal static bool MatchesSourceWorksheet(string workbookName, string workbookFullName,
        string qualifier, string sheetName)
    {
        if (qualifier.StartsWith('\'') && qualifier.EndsWith('\''))
            qualifier = qualifier[1..^1].Replace("''", "'", StringComparison.Ordinal);

        int workbookStart = qualifier.LastIndexOf('[');
        int workbookEnd = qualifier.IndexOf(']', workbookStart + 1);
        if (workbookStart >= 0 || workbookEnd >= 0)
        {
            if (workbookStart < 0 || workbookEnd <= workbookStart ||
                !string.Equals(qualifier[(workbookStart + 1)..workbookEnd], workbookName, StringComparison.OrdinalIgnoreCase))
                return false;
            if (workbookStart > 0 &&
                !string.Equals(qualifier[..workbookStart] + workbookName, workbookFullName, StringComparison.OrdinalIgnoreCase))
                return false;
            qualifier = qualifier[(workbookEnd + 1)..];
        }
        return string.Equals(qualifier, sheetName, StringComparison.OrdinalIgnoreCase);
    }
}
