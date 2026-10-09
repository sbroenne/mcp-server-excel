using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.PowerQuery;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

public partial class PivotTableCommands
{
    /// <inheritdoc/>
    public PivotConnectionResult GetConnection(IExcelBatch batch, string sheetName, string pivotTableName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(pivotTableName);
        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;
            Excel.PivotCache? cache = null;
            try
            {
                pivot = FindPivotOnSheet(ctx.Book, sheetName, pivotTableName, ct);
                cache = pivot.PivotCache();
                return ReadPivotConnection(ctx.Book, sheetName, pivot, cache, batch.WorkbookPath, ct);
            }
            finally
            {
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <inheritdoc/>
    public PivotConnectionResult SetConnection(IExcelBatch batch, string sheetName, string pivotTableName,
        string connectionName, TimeSpan? timeout = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(pivotTableName);
        ArgumentException.ThrowIfNullOrWhiteSpace(connectionName);
        using var timeoutCts = new CancellationTokenSource(timeout ?? TimeSpan.FromMinutes(5));
        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;
            Excel.PivotCache? cache = null;
            Excel.Worksheet? sheet = null;
            Excel.WorkbookConnection? current = null;
            Excel.WorkbookConnection? target = null;
            Excel.OLEDBConnection? oleDb = null;
            string phase = "validating the selected PivotTable";
            try
            {
                pivot = FindPivotOnSheet(ctx.Book, sheetName, pivotTableName, ct);
                cache = pivot.PivotCache();
                if (cache.SourceType != Excel.XlPivotTableSourceType.xlExternal)
                    throw new NotSupportedException("set-connection requires an external PivotTable. Use set-source for worksheet-backed PivotTables.");
                current = cache.WorkbookConnection;
                if (current is null || current.Type == Excel.XlConnectionType.xlConnectionTypeMODEL)
                    throw new NotSupportedException("set-connection cannot replace the workbook Data Model connection.");
                target = PowerQueryHelpers.FindConnectionByExactName(ctx.Book, connectionName);
                if (target is null)
                    throw new InvalidOperationException($"Connection '{connectionName}' not found.");
                if (target.Type is not (Excel.XlConnectionType.xlConnectionTypeOLEDB or Excel.XlConnectionType.xlConnectionTypeODBC) ||
                    target.Type != current.Type || PowerQueryHelpers.IsPowerQueryConnection(target))
                    throw new NotSupportedException("The target must be an ordinary external connection of the same OLEDB/ODBC type; workbook Data Model and Power Query targets are unsupported.");
                if (target.Type == Excel.XlConnectionType.xlConnectionTypeOLEDB)
                {
                    oleDb = target.OLEDBConnection;
                    if (oleDb.OLAP != cache.OLAP)
                        throw new NotSupportedException("The target connection must keep the PivotTable's existing OLAP mode.");
                }
                if (string.Equals(current.Name, target.Name, StringComparison.OrdinalIgnoreCase))
                    return ReadPivotConnection(ctx.Book, sheetName, pivot, cache, batch.WorkbookPath, ct);
                sheet = (Excel.Worksheet)pivot.Parent;
                if (sheet.ProtectContents)
                    throw new InvalidOperationException("Unprotect the PivotTable worksheet before changing its connection.");
                var before = ReadPivotConnection(ctx.Book, sheetName, pivot, cache, batch.WorkbookPath, ct);
                if (before.SharedPivotTables.Count > 1)
                    throw new InvalidOperationException("Other PivotTables share this cache. Changing their connection implicitly is unsupported; inspect get-connection before proceeding.");
                if (before.ConnectedSlicerCaches.Count > 0)
                    throw new InvalidOperationException($"Disconnect this PivotTable from slicer/timeline caches before changing its connection: {string.Join(", ", before.ConnectedSlicerCaches)}. Connected controls are not rebuilt.");
                ct.ThrowIfCancellationRequested();
                phase = "changing the connection";
                pivot.ChangeConnection(target);
                ComUtilities.Release(ref cache);
                phase = "reading the changed connection";
                cache = pivot.PivotCache();
                var result = ReadPivotConnection(ctx.Book, sheetName, pivot, cache, batch.WorkbookPath, ct);
                if (!string.Equals(result.ConnectionName, target.Name, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Excel did not apply the requested connection. Inspect get-connection before retrying; no changes have been rolled back.");
                return result;
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException(
                    $"Excel failed while {phase}. The connection may have changed; inspect get-connection before retrying. No changes have been rolled back.", ex);
            }
            finally
            {
                ComUtilities.Release(ref oleDb);
                ComUtilities.Release(ref target);
                ComUtilities.Release(ref current);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref pivot);
            }
        }, timeoutCts.Token);
    }

    private static Excel.PivotTable FindPivotOnSheet(Excel.Workbook book, string sheetName,
        string pivotTableName, CancellationToken ct)
    {
        Excel.Worksheet? sheet = null;
        Excel.PivotTables? pivots = null;
        try
        {
            sheet = ComUtilities.FindSheet(book, sheetName);
            if (sheet is null)
                throw new InvalidOperationException($"Worksheet '{sheetName}' not found.");
            pivots = sheet.PivotTables();
            for (int index = 1; index <= pivots.Count; index++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.PivotTable? pivot = null;
                try
                {
                    pivot = pivots.Item(index);
                    if (string.Equals(pivot.Name, pivotTableName, StringComparison.OrdinalIgnoreCase))
                    {
                        var found = pivot;
                        pivot = null;
                        return found;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref pivot);
                }
            }
            throw new InvalidOperationException($"PivotTable '{pivotTableName}' not found on worksheet '{sheetName}'.");
        }
        finally
        {
            ComUtilities.Release(ref pivots);
            ComUtilities.Release(ref sheet);
        }
    }

    private static PivotConnectionResult ReadPivotConnection(Excel.Workbook book, string sheetName,
        Excel.PivotTable pivot, Excel.PivotCache cache, string filePath, CancellationToken ct)
    {
        Excel.WorkbookConnection? connection = null;
        Excel.Sheets? sheets = null;
        Excel.SlicerCaches? slicerCaches = null;
        try
        {
            if (cache.SourceType == Excel.XlPivotTableSourceType.xlExternal)
                connection = cache.WorkbookConnection;
            var result = new PivotConnectionResult
            {
                Success = true,
                FilePath = filePath,
                SheetName = sheetName,
                PivotTableName = pivot.Name,
                ConnectionName = connection?.Name,
                CacheIndex = pivot.CacheIndex,
                IsOlap = cache.OLAP,
                IsDataModel = connection?.Type == Excel.XlConnectionType.xlConnectionTypeMODEL
            };
            result.IsExternal = connection is not null && !result.IsDataModel;
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
                        Excel.PivotTable? other = null;
                        try
                        {
                            other = pivots.Item(position);
                            if (other.CacheIndex == result.CacheIndex)
                                result.SharedPivotTables.Add(new PivotTableIdentity { SheetName = sheet.Name, PivotTableName = other.Name });
                        }
                        finally
                        {
                            ComUtilities.Release(ref other);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref pivots);
                    ComUtilities.Release(ref sheet);
                }
            }
            slicerCaches = book.SlicerCaches;
            for (int index = 1; index <= slicerCaches.Count; index++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.SlicerCache? slicerCache = null;
                Excel.SlicerPivotTables? pivots = null;
                try
                {
                    slicerCache = slicerCaches[index];
                    if (slicerCache.List)
                        continue;
                    pivots = slicerCache.PivotTables;
                    for (int position = 1; position <= pivots.Count; position++)
                    {
                        ct.ThrowIfCancellationRequested();
                        Excel.PivotTable? other = null;
                        Excel.Worksheet? sheet = null;
                        try
                        {
                            other = pivots[position];
                            sheet = (Excel.Worksheet)other.Parent;
                            if (string.Equals(other.Name, result.PivotTableName, StringComparison.OrdinalIgnoreCase) &&
                                string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                            {
                                result.ConnectedSlicerCaches.Add(slicerCache.Name);
                                break;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref sheet);
                            ComUtilities.Release(ref other);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref pivots);
                    ComUtilities.Release(ref slicerCache);
                }
            }
            return result;
        }
        finally
        {
            ComUtilities.Release(ref slicerCaches);
            ComUtilities.Release(ref sheets);
            ComUtilities.Release(ref connection);
        }
    }
}
