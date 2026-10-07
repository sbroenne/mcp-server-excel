using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Slicer;

/// <summary>
/// Checks shared by slicer and timeline creation. They run before Excel creates a slicer cache,
/// because a cache created for a request that then fails stays in the workbook.
/// </summary>
internal static class SlicerPlacement
{
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
        Excel.Workbook book, string sheetName, string position)
    {
        Excel.Worksheet? sheet;
        Excel.Sheets? worksheets = null;
        try
        {
            worksheets = book.Worksheets;
            sheet = (Excel.Worksheet)worksheets[sheetName];
        }
        catch (COMException ex)
        {
            throw new ArgumentException($"Worksheet '{sheetName}' not found.", ex);
        }
        finally
        {
            ComUtilities.Release(ref worksheets);
        }

        try
        {
            return (sheet, sheet.Range[position]);
        }
        catch (COMException ex)
        {
            ComUtilities.Release(ref sheet);
            throw new ArgumentException(
                $"Position '{position}' is not a valid cell reference on sheet '{sheetName}'.", ex);
        }
    }

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
