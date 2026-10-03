using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

public partial class PivotTableCommands
{
    /// <inheritdoc/>
    public PivotSourceResult GetSource(IExcelBatch batch, string pivotTableName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(pivotTableName);
        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;
            Excel.PivotCache? cache = null;
            try
            {
                pivot = (Excel.PivotTable)FindPivotTable(ctx.Book, pivotTableName);
                cache = pivot.PivotCache();
                RequireWorksheetPivotCache(cache);
                return ReadPivotSource(ctx.Book, pivot, cache, batch.WorkbookPath, ct);
            }
            finally
            {
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <inheritdoc/>
    public PivotSourceResult SetSource(IExcelBatch batch, string pivotTableName, string sourceSheetName,
        string? sourceRangeAddress = null, string? tableName = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(pivotTableName);
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceSheetName);
        if ((sourceRangeAddress is null) == (tableName is null))
            throw new ArgumentException("Specify exactly one sourceRangeAddress or tableName.");
        if (sourceRangeAddress is not null)
            ArgumentException.ThrowIfNullOrWhiteSpace(sourceRangeAddress);
        if (tableName is not null)
            ArgumentException.ThrowIfNullOrWhiteSpace(tableName);
        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;
            Excel.PivotCache? oldCache = null;
            Excel.Worksheet? destination = null;
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sourceSheet = null;
            Excel.ListObjects? tables = null;
            Excel.ListObject? table = null;
            Excel.Range? source = null;
            Excel.Range? rows = null;
            Excel.Range? header = null;
            Excel.Areas? areas = null;
            Excel.PivotTables? sourcePivots = null;
            Excel.PivotFields? fields = null;
            Excel.PivotCaches? caches = null;
            Excel.PivotCache? newCache = null;
            string phase = "validating the source";
            try
            {
                pivot = (Excel.PivotTable)FindPivotTable(ctx.Book, pivotTableName);
                oldCache = pivot.PivotCache();
                RequireWorksheetPivotCache(oldCache);
                if (!oldCache.EnableRefresh)
                    throw new InvalidOperationException("Enable cache refresh before replacing and refreshing the source.");
                destination = (Excel.Worksheet)pivot.Parent;
                if (destination.ProtectContents)
                    throw new InvalidOperationException("Unprotect the PivotTable worksheet before replacing its source.");
                var connected = ConnectedPivotSlicerCaches(ctx.Book, pivot.Name, ct);
                if (connected.Count > 0)
                    throw new InvalidOperationException($"Disconnect this PivotTable from slicer/timeline caches before source replacement: {string.Join(", ", connected)}. Their shared caches are not rebuilt.");
                sheets = ctx.Book.Worksheets;
                sourceSheet = (Excel.Worksheet)sheets[sourceSheetName];
                if (tableName is not null)
                {
                    tables = sourceSheet.ListObjects;
                    table = tables[tableName];
                    if (!table.ShowHeaders)
                        throw new ArgumentException("The source table must have visible field headers.");
                    source = table.Range;
                    // Table totals are not source records; the structured name handles their exclusion.
                }
                else
                    source = sourceSheet.Range[sourceRangeAddress!];
                areas = source.Areas;
                if (areas.Count != 1)
                    throw new ArgumentException("The source must be one contiguous range.");
                rows = source.Rows;
                if (rows.Count < 2)
                    throw new ArgumentException("The source requires headers and at least one record.");
                header = (Excel.Range)rows[1];
                var values = ExcelValueNormalizer.Normalize(header.Value2).Values[0];
                var headers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var value in values)
                {
                    ct.ThrowIfCancellationRequested();
                    if (value is not string text || string.IsNullOrWhiteSpace(text) || !headers.Add(text))
                        throw new ArgumentException("Source headers must be distinct, nonempty text field names.");
                }
                fields = (Excel.PivotFields)pivot.PivotFields();
                var sourceFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (int index = 1; index <= fields.Count; index++)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.PivotField? field = null;
                    try
                    {
                        field = fields.Item(index);
                        if (field.IsCalculated)
                            continue;
                        sourceFields.Add(field.SourceName);
                    }
                    finally
                    {
                        ComUtilities.Release(ref field);
                    }
                }
                if (!sourceFields.SetEquals(headers))
                    throw new ArgumentException("Replacement source must keep exactly the existing source field names. Field schema changes are not silently rebuilt.");
                sourcePivots = sourceSheet.PivotTables();
                for (int index = 1; index <= sourcePivots.Count; index++)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.PivotTable? other = null;
                    Excel.Range? extent = null;
                    Excel.Range? overlap = null;
                    try
                    {
                        other = sourcePivots.Item(index);
                        extent = other.TableRange2;
                        overlap = ctx.App.Intersect(source, extent);
                        if (overlap is not null)
                            throw new ArgumentException("The replacement source cannot overlap any PivotTable output.");
                    }
                    finally
                    {
                        ComUtilities.Release(ref overlap);
                        ComUtilities.Release(ref extent);
                        ComUtilities.Release(ref other);
                    }
                }
                string reference = table is not null ? table.Name :
                    source.Address[true, true, Excel.XlReferenceStyle.xlR1C1, true];
                caches = ctx.Book.PivotCaches();
                ct.ThrowIfCancellationRequested();
                phase = "creating the replacement cache";
                newCache = caches.Create(Excel.XlPivotTableSourceType.xlDatabase, reference, oldCache.Version);
                phase = "preserving refresh-on-open";
                if (newCache.RefreshOnFileOpen != oldCache.RefreshOnFileOpen)
                    newCache.RefreshOnFileOpen = oldCache.RefreshOnFileOpen;
                phase = "preserving missing-item retention";
                if (newCache.MissingItemsLimit != oldCache.MissingItemsLimit)
                    newCache.MissingItemsLimit = oldCache.MissingItemsLimit;
                phase = "preserving cache optimization";
                if (newCache.OptimizeCache != oldCache.OptimizeCache)
                    newCache.OptimizeCache = oldCache.OptimizeCache;
                phase = "changing the selected PivotTable cache";
                pivot.ChangePivotCache(newCache);
                phase = "refreshing the new source";
                if (!pivot.RefreshTable())
                    throw new InvalidOperationException("Excel did not refresh the replaced source. Read get-source before retrying; failed writes do not promise rollback.");
                phase = "reading the replaced source";
                return ReadPivotSource(ctx.Book, pivot, newCache, batch.WorkbookPath, ct);
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException(
                    $"Excel failed while {phase}. Read get-source before retrying; a failed write does not promise rollback.", ex);
            }
            finally
            {
                ComUtilities.Release(ref newCache);
                ComUtilities.Release(ref caches);
                ComUtilities.Release(ref fields);
                ComUtilities.Release(ref sourcePivots);
                ComUtilities.Release(ref areas);
                ComUtilities.Release(ref header);
                ComUtilities.Release(ref rows);
                ComUtilities.Release(ref source);
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref tables);
                ComUtilities.Release(ref sourceSheet);
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref destination);
                ComUtilities.Release(ref oldCache);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    private static void RequireWorksheetPivotCache(Excel.PivotCache cache)
    {
        if (cache.OLAP || cache.SourceType != Excel.XlPivotTableSourceType.xlDatabase)
            throw new NotSupportedException("Source inspection/replacement supports worksheet-backed regular PivotTables only. External and OLAP/Data Model source mutations are unsupported.");
    }

    private static PivotSourceResult ReadPivotSource(Excel.Workbook book, Excel.PivotTable pivot,
        Excel.PivotCache cache, string filePath, CancellationToken ct) => new()
        {
            Success = true,
            FilePath = filePath,
            PivotTableName = pivot.Name,
            CacheIndex = pivot.CacheIndex,
            SourceData = Convert.ToString(cache.SourceData, CultureInfo.InvariantCulture) ?? string.Empty,
            RecordCount = cache.RecordCount,
            SharedPivotTables = ReadSharedPivotTables(book, pivot.CacheIndex, ct),
            ConnectedSlicerCaches = ConnectedPivotSlicerCaches(book, pivot.Name, ct)
        };

    private static List<string> ReadSharedPivotTables(Excel.Workbook book, int cacheIndex, CancellationToken ct)
    {
        List<string> names = [];
        Excel.Sheets? sheets = null;
        try
        {
            sheets = book.Worksheets;
            for (int index = 1; index <= sheets.Count; index++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.Worksheet? sheet = null;
                Excel.PivotTables? pivots = null;
                try
                {
                    sheet = (Excel.Worksheet)sheets[index];
                    pivots = sheet.PivotTables();
                    for (int position = 1; position <= pivots.Count; position++)
                    {
                        ct.ThrowIfCancellationRequested();
                        Excel.PivotTable? pivot = null;
                        try
                        {
                            pivot = pivots.Item(position);
                            if (pivot.CacheIndex == cacheIndex)
                                names.Add(pivot.Name);
                        }
                        finally
                        {
                            ComUtilities.Release(ref pivot);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref pivots);
                    ComUtilities.Release(ref sheet);
                }
            }
            return names;
        }
        finally
        {
            ComUtilities.Release(ref sheets);
        }
    }

    private static List<string> ConnectedPivotSlicerCaches(Excel.Workbook book, string name, CancellationToken ct)
    {
        List<string> names = [];
        Excel.SlicerCaches? caches = null;
        try
        {
            caches = book.SlicerCaches;
            for (int index = 1; index <= caches.Count; index++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.SlicerCache? cache = null;
                Excel.SlicerPivotTables? pivots = null;
                try
                {
                    cache = caches[index];
                    if (cache.List)
                        continue;
                    pivots = cache.PivotTables;
                    for (int position = 1; position <= pivots.Count; position++)
                    {
                        ct.ThrowIfCancellationRequested();
                        Excel.PivotTable? pivot = null;
                        try
                        {
                            pivot = pivots[position];
                            if (string.Equals(pivot.Name, name, StringComparison.OrdinalIgnoreCase))
                            {
                                names.Add(cache.Name);
                                break;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref pivot);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref pivots);
                    ComUtilities.Release(ref cache);
                }
            }
            return names;
        }
        finally
        {
            ComUtilities.Release(ref caches);
        }
    }
}
