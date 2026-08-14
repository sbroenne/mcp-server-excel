using Excel = Microsoft.Office.Interop.Excel;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable and PivotCache configuration.
/// </summary>
public partial class PivotTableCommands
{
    /// <inheritdoc />
    public PivotCacheOptionsResult GetCacheOptions(IExcelBatch batch, string pivotTableName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;
            Excel.PivotCache? cache = null;
            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                cache = pivot.PivotCache();
                return ReadCacheOptions(pivot, cache, pivotTableName, batch.WorkbookPath);
            }
            finally
            {
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <inheritdoc />
    public PivotCacheOptionsResult SetCacheOptions(
        IExcelBatch batch,
        string pivotTableName,
        bool? enableRefresh = null,
        bool? refreshOnFileOpen = null,
        PivotMissingItemsLimit? missingItemsLimit = null,
        bool? optimizeCache = null,
        bool? saveSourceData = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;
            Excel.PivotCache? cache = null;
            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                cache = pivot.PivotCache();
                var isOlap = cache.OLAP;
                var isExternal = cache.SourceType == Excel.XlPivotTableSourceType.xlExternal;

                if (missingItemsLimit.HasValue && isOlap)
                {
                    throw new InvalidOperationException(
                        "Missing-item retention is not available for OLAP/Data Model PivotCaches. " +
                        "Configure retained members in the external model or source.");
                }

                if (enableRefresh.HasValue)
                {
                    cache.EnableRefresh = enableRefresh.Value;
                }

                if (refreshOnFileOpen.HasValue)
                {
                    cache.RefreshOnFileOpen = refreshOnFileOpen.Value;
                }

                if (missingItemsLimit.HasValue)
                {
                    cache.MissingItemsLimit = (Excel.XlPivotTableMissingItems)missingItemsLimit.Value;
                }

                if (optimizeCache.HasValue && isExternal)
                {
                    if (optimizeCache.Value != cache.OptimizeCache)
                    {
                        throw new InvalidOperationException(
                            "OptimizeCache is read-only for external OLE DB/OLAP PivotCaches.");
                    }
                }
                else if (optimizeCache.HasValue)
                {
                    cache.OptimizeCache = optimizeCache.Value;
                }

                if (saveSourceData == true && isOlap)
                {
                    throw new InvalidOperationException(
                        "OLAP/Data Model PivotTables cannot save source records in the workbook.");
                }
                else if (saveSourceData.HasValue && !isOlap)
                {
                    pivot.SaveData = saveSourceData.Value;
                }

                return ReadCacheOptions(pivot, cache, pivotTableName, batch.WorkbookPath);
            }
            finally
            {
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    private static PivotCacheOptionsResult ReadCacheOptions(
        Excel.PivotTable pivot,
        Excel.PivotCache cache,
        string pivotTableName,
        string filePath)
    {
        var isOlap = cache.OLAP;
        return new PivotCacheOptionsResult
        {
            Success = true,
            PivotTableName = pivotTableName,
            EnableRefresh = cache.EnableRefresh,
            RefreshOnFileOpen = cache.RefreshOnFileOpen,
            MissingItemsLimit = isOlap
                ? null
                : (PivotMissingItemsLimit)cache.MissingItemsLimit,
            OptimizeCache = cache.OptimizeCache,
            SaveSourceData = pivot.SaveData,
            IsOlap = isOlap,
            FilePath = filePath
        };
    }
}
