using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Slicer;

/// <summary>
/// Checks shared by slicer and timeline creation. They run before Excel creates a slicer cache,
/// because a cache created for a request that then fails stays in the workbook.
/// </summary>
internal static class SlicerPlacement
{
    private const int ExcelMaxRows = 1_048_576;
    private const int ExcelMaxColumns = 16_384;
    private static readonly Regex CellAddress = new(
        @"^\$?(?<column>[A-Z]{1,3})\$?(?<row>\d{1,7})$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    /// <summary>
    /// Rejects a name already used by any slicer or timeline in the workbook.
    /// </summary>
    internal static void ValidateNewControlName(Excel.SlicerCaches caches, string name, CancellationToken ct)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        for (int index = 1; index <= caches.Count; index++)
        {
            ct.ThrowIfCancellationRequested();
            Excel.SlicerCache? cache = null;
            Excel.Slicers? slicers = null;
            try
            {
                cache = caches[index];
                slicers = cache.Slicers;
                for (int item = 1; item <= slicers.Count; item++)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.Slicer? slicer = null;
                    try
                    {
                        slicer = slicers[item];
                        if (string.Equals(slicer.Name, name, StringComparison.OrdinalIgnoreCase))
                            throw new ArgumentException($"Slicer or timeline '{name}' already exists.");
                    }
                    finally
                    {
                        ComUtilities.Release(ref slicer);
                    }
                }
            }
            finally
            {
                ComUtilities.Release(ref slicers);
                ComUtilities.Release(ref cache);
            }
        }
    }

    /// <summary>
    /// Finds the destination sheet and anchor cell. The caller releases both.
    /// </summary>
    internal static (Excel.Worksheet Sheet, Excel.Range Anchor) ResolveDestination(
        Excel.Workbook book, string sheetName, string position, CancellationToken ct)
    {
        var (column, row) = ParsePosition(sheetName, position);
        var sheet = FindWorksheet(book, sheetName, ct);
        Excel.Range? cells = null;
        try
        {
            cells = sheet.Cells;
            Excel.Range anchor = cells[row, column];
            return (sheet, anchor);
        }
        catch
        {
            ComUtilities.Release(ref sheet);
            throw;
        }
        finally
        {
            ComUtilities.Release(ref cells);
        }
    }

    private static Excel.Worksheet FindWorksheet(
        Excel.Workbook book, string sheetName, CancellationToken ct)
    {
        Excel.Sheets? worksheets = null;
        Excel.Worksheet? candidate = null;
        try
        {
            worksheets = book.Worksheets;
            for (int index = 1; index <= worksheets.Count; index++)
            {
                ct.ThrowIfCancellationRequested();
                candidate = worksheets[index];
                if (string.Equals(candidate.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    var result = candidate;
                    candidate = null;
                    return result;
                }

                ComUtilities.Release(ref candidate);
            }

            throw new ArgumentException($"Worksheet '{sheetName}' not found.");
        }
        finally
        {
            ComUtilities.Release(ref candidate);
            ComUtilities.Release(ref worksheets);
        }
    }

    private static (int Column, int Row) ParsePosition(string sheetName, string position)
    {
        if (string.IsNullOrWhiteSpace(position))
        {
            throw InvalidPosition(sheetName, position);
        }

        var match = CellAddress.Match(position);
        if (!match.Success)
        {
            throw InvalidPosition(sheetName, position);
        }

        int column = 0;
        foreach (char letter in match.Groups["column"].Value)
        {
            column = (column * 26) + (char.ToUpperInvariant(letter) - 'A' + 1);
        }

        if (!int.TryParse(match.Groups["row"].Value, out int row) ||
            column is < 1 or > ExcelMaxColumns ||
            row is < 1 or > ExcelMaxRows)
        {
            throw InvalidPosition(sheetName, position);
        }

        return (column, row);
    }

    private static ArgumentException InvalidPosition(string sheetName, string position) =>
        new($"Position '{position}' is not a valid cell reference on sheet '{sheetName}'.", nameof(position));

    /// <summary>
    /// Builds the error for a slicer that Excel failed to add after it created a new slicer cache.
    /// </summary>
    internal static InvalidOperationException LeftoverCache(Excel.SlicerCache cache, string fieldName, Exception cause)
    {
        string? cacheName = null;
        try
        {
            cacheName = cache.Name;
        }
        catch (COMException)
        {
            // The name is only used in the message; the original failure matters more.
        }

        var created = cacheName is null ? $"a slicer cache for field '{fieldName}'" : $"slicer cache '{cacheName}'";
        var subject = cacheName is null ? "That slicer cache" : $"The slicer cache '{cacheName}'";
        return CreatedObjectFailure.CreateDescribed(
            "Adding the slicer",
            created,
            $"{subject} remains in the workbook without a slicer; creating another slicer for field '{fieldName}' reuses it.",
            cause);
    }
}
