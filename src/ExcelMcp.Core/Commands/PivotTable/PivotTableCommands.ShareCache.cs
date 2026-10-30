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
            bool writeStarted = false;
            string phase = "finding the reports";
            int previousIndex = 0;
            List<PivotTableIdentity> previousUsers = [];
            try
            {
                pivot = FindPivotOnSheet(ctx.Book, sheetName, pivotTableName, ct);
                target = FindPivotOnSheet(ctx.Book, shareCacheWithSheet, shareCacheWith, ct);
                oldCache = pivot.PivotCache();
                previousIndex = pivot.CacheIndex;
                previousUsers = ReadPivotConnection(ctx.Book, sheetName, pivot, oldCache,
                    batch.WorkbookPath, ct).SharedPivotTables;
                if (previousIndex == target.CacheIndex)
                    return ReadCacheSharingResult(ctx.Book, pivot, target, sheetName,
                        shareCacheWithSheet, previousIndex, previousUsers, false, batch.WorkbookPath, ct);

                phase = "assigning the selected report's cache";
                writeStarted = true;
                pivot.CacheIndex = target.CacheIndex;
                phase = "reading the actual cache";
                var result = ReadCacheSharingResult(ctx.Book, pivot, target, sheetName,
                    shareCacheWithSheet, previousIndex, previousUsers, true, batch.WorkbookPath, ct);
                if (result.CacheIndex != result.TargetCacheIndex)
                    throw new InvalidOperationException("Excel did not assign the chosen cache.");
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
                throw new InvalidOperationException($"Excel failed while {phase}, before cache assignment.", ex);
            }
            finally
            {
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
            $"Excel failed while {phase}: {failure.Message}. No changes were undone. Actual partial state: {state}", cause);
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
}
