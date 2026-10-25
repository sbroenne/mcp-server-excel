using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

public partial class PivotTableCommands
{
    /// <inheritdoc/>
    public PivotCacheSharingResult ShareCache(IExcelBatch batch, string sheetName, string pivotTableName,
        string shareCacheWith, string shareCacheWithSheet)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(pivotTableName);
        ArgumentException.ThrowIfNullOrWhiteSpace(shareCacheWith);
        ArgumentException.ThrowIfNullOrWhiteSpace(shareCacheWithSheet);
        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;
            Excel.PivotTable? target = null;
            Excel.PivotCache? oldCache = null;
            Excel.PivotCache? targetCache = null;
            Excel.Worksheet? sheet = null;
            bool writeStarted = false;
            string phase = "checking the reports";
            int previousIndex = 0;
            List<PivotTableIdentity> previousUsers = [];
            try
            {
                pivot = FindPivotOnSheet(ctx.Book, sheetName, pivotTableName, ct);
                target = FindPivotOnSheet(ctx.Book, shareCacheWithSheet, shareCacheWith, ct);
                oldCache = pivot.PivotCache();
                targetCache = target.PivotCache();
                RequireWorksheetPivotCache(oldCache);
                RequireWorksheetPivotCache(targetCache);
                previousIndex = pivot.CacheIndex;
                var before = ReadPivotConnection(ctx.Book, sheetName, pivot, oldCache, batch.WorkbookPath, ct);
                previousUsers = before.SharedPivotTables;
                if (previousIndex == target.CacheIndex)
                    return ReadCacheSharingResult(ctx.Book, pivot, target, sheetName,
                        shareCacheWithSheet, previousIndex, previousUsers, false, batch.WorkbookPath, ct);
                sheet = (Excel.Worksheet)pivot.Parent;
                if (sheet.ProtectContents)
                    throw new InvalidOperationException("Unprotect the moved report's worksheet before sharing its cache.");
                if (before.ConnectedSlicerCaches.Count > 0)
                    throw new InvalidOperationException(
                        $"Disconnect the moved report from slicer/timeline caches first: {string.Join(", ", before.ConnectedSlicerCaches)}. Controls are not rebuilt or reconnected.");
                phase = "checking source identity";
                var source = ReadCacheSharingSource(ctx.Book, ctx.App, oldCache, ct);
                if (source.Identity != ReadCacheSharingSource(ctx.Book, ctx.App, targetCache, ct).Identity)
                    throw new ArgumentException("Cannot share the cache: source definitions differ. Tables and fixed ranges are distinct even when their data matches.");
                phase = "checking fields and calculations";
                var oldFields = ReadCacheSharingFields(pivot, source.NumericFields, ct);
                var targetFields = ReadCacheSharingFields(target, source.NumericFields, ct);
                foreach (var field in oldFields)
                {
                    if (!targetFields.TryGetValue(field.Key, out var other) ||
                        field.Value != other)
                        throw new NotSupportedException(
                            $"Field '{field.Key}' has incompatible cached items, grouping, or a missing/conflicting calculated formula. No cache was changed.");
                }
                if (targetFields.Any(field => !oldFields.ContainsKey(field.Key) && !field.Value.StartsWith("formula:", StringComparison.Ordinal)))
                    throw new NotSupportedException("The caches must have the same source fields. No cache was changed.");
                phase = "checking the report settings";
                string reportState = ReadSharingReportState(pivot, batch.WorkbookPath, ct);
                ct.ThrowIfCancellationRequested();
                phase = "changing the selected report's cache";
                writeStarted = true;
                pivot.CacheIndex = target.CacheIndex;
                phase = "reading the actual cache";
                var result = ReadCacheSharingResult(ctx.Book, pivot, target, sheetName,
                    shareCacheWithSheet, previousIndex, previousUsers, true, batch.WorkbookPath, ct);
                if (result.CacheIndex != result.TargetCacheIndex)
                    throw new InvalidOperationException("Excel did not assign the chosen cache.");
                phase = "verifying preserved report settings";
                if (reportState != ReadSharingReportState(pivot, batch.WorkbookPath, ct))
                    throw new InvalidOperationException("Excel changed report settings during cache reassignment. Inspect the moved report; no changes were undone.");
                return result;
            }
            catch (Exception ex) when (writeStarted && ex is not OperationCanceledException)
            {
                throw DescribeCacheSharingFailure(ex, phase, () =>
                    ReadCacheSharingResult(ctx.Book, pivot!, target!, sheetName,
                        shareCacheWithSheet, previousIndex, previousUsers, true, batch.WorkbookPath, ct));
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException($"Excel failed while {phase}, before cache reassignment.", ex);
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref targetCache);
                ComUtilities.Release(ref oldCache);
                ComUtilities.Release(ref target);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    [SuppressMessage("COM", "EXL002", Justification = "A failed post-write state inspection is combined with the original failure and rethrown; neither failure becomes a successful or default result.")]
    internal static InvalidOperationException DescribeCacheSharingFailure(Exception failure, string phase,
        Func<PivotCacheSharingResult> inspect)
    {
        string state;
        Exception cause = failure;
        try
        {
            var partial = inspect();
            partial.Success = false;
            state = JsonSerializer.Serialize(partial, JsonSerializerOptions.Web);
        }
        catch (Exception readFailure) when (readFailure is not OperationCanceledException)
        {
            state = $"Actual cache state could not be read: {readFailure.Message}";
            cause = new AggregateException(failure, readFailure);
        }
        return new InvalidOperationException(
            $"Excel failed while {phase}. No changes were undone. Actual partial state: {state}", cause);
    }

    private static PivotCacheSharingResult ReadCacheSharingResult(Excel.Workbook book,
        Excel.PivotTable pivot, Excel.PivotTable target, string sheetName, string targetSheet,
        int previousIndex, List<PivotTableIdentity> previousUsers, bool changed, string filePath, CancellationToken ct)
    {
        Excel.PivotCache? cache = null;
        try
        {
            cache = pivot.PivotCache();
            var current = ReadPivotConnection(book, sheetName, pivot, cache, filePath, ct);
            bool moved = changed && current.CacheIndex == target.CacheIndex;
            return new PivotCacheSharingResult
            {
                Success = true,
                FilePath = filePath,
                SheetName = sheetName,
                PivotTableName = pivot.Name,
                ShareCacheWithSheet = targetSheet,
                ShareCacheWith = target.Name,
                PreviousCacheIndex = previousIndex,
                CacheIndex = current.CacheIndex,
                TargetCacheIndex = target.CacheIndex,
                Changed = moved,
                SharedPivotTables = current.SharedPivotTables,
                PreviousCacheUsers = previousUsers.Where(user => !moved ||
                    !current.SharedPivotTables.Any(shared => shared.SheetName == user.SheetName &&
                        shared.PivotTableName == user.PivotTableName)).ToList()
            };
        }
        finally
        {
            ComUtilities.Release(ref cache);
        }
    }

    private sealed record CacheSharingSource(string Identity, HashSet<string> NumericFields);

    private static CacheSharingSource ReadCacheSharingSource(Excel.Workbook book, Excel.Application app,
        Excel.PivotCache cache, CancellationToken ct)
    {
        if (cache.SourceData is not string reference)
            throw new NotSupportedException("The cache source must be a resolvable worksheet Table or fixed range.");
        int separator = reference.LastIndexOf('!');
        string address = separator < 0 ? reference : reference[(separator + 1)..];
        Excel.Sheets? sheets = null;
        try
        {
            sheets = book.Worksheets;
            for (int i = 1; i <= sheets.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.Worksheet? sheet = null;
                Excel.ListObjects? tables = null;
                Excel.Range? range = null;
                try
                {
                    sheet = (Excel.Worksheet)sheets[i];
                    if (separator >= 0 && !MatchesSourceWorksheet(book.Name, book.FullName, reference[..separator], sheet.Name))
                        continue;
                    tables = sheet.ListObjects;
                    for (int j = 1; j <= tables.Count; j++)
                    {
                        Excel.ListObject? table = null;
                        try
                        {
                            table = tables[j];
                            if (string.Equals(address, table.Name, StringComparison.OrdinalIgnoreCase))
                            {
                                range = table.Range;
                                return new CacheSharingSource($"table:{sheet.Name}:{table.Name}".ToUpperInvariant(),
                                    ReadNumericSourceFields(app, range, ct));
                            }
                        }
                        finally { ComUtilities.Release(ref table); }
                    }
                    if (separator < 0)
                        continue;
                    if (IsNativeR1C1SourceAddress(address))
                        address = Convert.ToString((object)app.ConvertFormula(address, Excel.XlReferenceStyle.xlR1C1,
                            Excel.XlReferenceStyle.xlA1, Excel.XlReferenceType.xlAbsolute), CultureInfo.InvariantCulture)!;
                    else if (!IsFixedA1SourceAddress(address))
                        break;
                    range = sheet.Range[address];
                    return new CacheSharingSource($"range:{sheet.Name}:{range.Address[true, true, Excel.XlReferenceStyle.xlA1]}".ToUpperInvariant(),
                        ReadNumericSourceFields(app, range, ct));
                }
                finally
                {
                    ComUtilities.Release(ref range);
                    ComUtilities.Release(ref tables);
                    ComUtilities.Release(ref sheet);
                }
            }
            throw new NotSupportedException("The cache source must be a resolvable worksheet Table or fixed range; named, dynamic, and external definitions are not supported.");
        }
        finally { ComUtilities.Release(ref sheets); }
    }

    private static HashSet<string> ReadNumericSourceFields(Excel.Application app, Excel.Range source, CancellationToken ct)
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        Excel.WorksheetFunction? functions = null;
        Excel.Range? columns = null;
        Excel.Range? rows = null;
        Excel.Range? cells = null;
        try
        {
            functions = app.WorksheetFunction;
            columns = source.Columns;
            rows = source.Rows;
            cells = source.Cells;
            if (rows.Count < 2)
                throw new NotSupportedException("The cache source requires headers and at least one record.");
            for (int i = 1; i <= columns.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.Range? header = null;
                Excel.Range? offset = null;
                Excel.Range? body = null;
                try
                {
                    header = (Excel.Range)cells[1, i];
                    offset = header.Offset[1, 0];
                    body = offset.Resize[rows.Count - 1, 1];
                    double count = functions.Count(body);
                    if (count > 0 && count == functions.CountA(body))
                        result.Add(Convert.ToString((object)header.Value2, CultureInfo.InvariantCulture)!);
                }
                finally
                {
                    ComUtilities.Release(ref body);
                    ComUtilities.Release(ref offset);
                    ComUtilities.Release(ref header);
                }
            }
            return result;
        }
        finally
        {
            ComUtilities.Release(ref cells);
            ComUtilities.Release(ref rows);
            ComUtilities.Release(ref columns);
            ComUtilities.Release(ref functions);
        }
    }

    private static Dictionary<string, string> ReadCacheSharingFields(Excel.PivotTable pivot,
        HashSet<string> numericSourceFields, CancellationToken ct)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        Excel.PivotFields? fields = null;
        try
        {
            fields = (Excel.PivotFields)pivot.PivotFields();
            for (int i = 1; i <= fields.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.PivotField? field = null;
                Excel.CalculatedItems? calculated = null;
                Excel.PivotItems? items = null;
                try
                {
                    field = fields.Item(i);
                    if (PivotTableHelpers.IsValuesPlaceholderField(field, field.SourceName))
                        continue;
                    if (field.GroupLevel > 1)
                        throw new NotSupportedException($"Grouped field '{field.Name}' is unsupported for cache sharing. No cache was changed.");
                    if (field.IsCalculated)
                    {
                        result[field.SourceName] = "formula:" + field.Formula;
                        continue;
                    }
                    if (numericSourceFields.Contains(field.SourceName) && field.DataType == Excel.XlPivotFieldDataType.xlText)
                        throw new NotSupportedException(
                            $"Grouped field or stale numeric schema in '{field.Name}' is unsupported. No cache was changed.");
                    calculated = field.CalculatedItems();
                    if (calculated.Count != 0)
                        throw new NotSupportedException($"Calculated items in '{field.Name}' are unsupported for cache sharing. No cache was changed.");
                    items = (Excel.PivotItems)field.PivotItems();
                    List<string> members = [];
                    for (int j = 1; j <= items.Count; j++)
                    {
                        ct.ThrowIfCancellationRequested();
                        Excel.PivotItem? item = null;
                        try
                        {
                            item = items.Item(j);
                            string member = Convert.ToString((object)item.SourceNameStandard, CultureInfo.InvariantCulture)!;
                            if (field.DataType == Excel.XlPivotFieldDataType.xlNumber &&
                                member != "(blank)" &&
                                !double.TryParse(member, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out _))
                                throw new NotSupportedException(
                                    $"Grouped field or unsupported numeric member in '{field.Name}': '{member}'. No cache was changed.");
                            members.Add(member);
                        }
                        finally { ComUtilities.Release(ref item); }
                    }
                    result[field.SourceName] = $"{field.DataType}:" + JsonSerializer.Serialize(members.Order(StringComparer.Ordinal));
                }
                finally
                {
                    ComUtilities.Release(ref items);
                    ComUtilities.Release(ref calculated);
                    ComUtilities.Release(ref field);
                }
            }
            return result;
        }
        finally { ComUtilities.Release(ref fields); }
    }
}
