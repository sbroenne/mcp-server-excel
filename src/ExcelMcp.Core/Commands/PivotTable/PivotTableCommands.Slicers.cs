using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Slicer;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable slicer operations (CreateSlicer, ListSlicers, SetSlicerSelection, DeleteSlicer)
/// </summary>
public partial class PivotTableCommands
{
    /// <summary>
    /// Creates a slicer for a PivotTable field
    /// </summary>
    public SlicerResult CreateSlicer(IExcelBatch batch, string pivotTableName,
        string fieldName, string slicerName, string destinationSheet, string position)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;
            dynamic? slicerCaches = null;
            dynamic? slicerCache = null;
            dynamic? slicers = null;
            dynamic? slicer = null;
            dynamic? destSheet = null;
            dynamic? destRange = null;
            Excel.Sheets? sheets = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                slicerCaches = ctx.Book.SlicerCaches;

                // Check if a SlicerCache already exists for this field+PivotTable
                // If so, we add a new visual Slicer to the existing cache
                slicerCache = FindExistingSlicerCache(slicerCaches, pivot, fieldName, ct);

                if (slicerCache == null)
                {
                    // Create new SlicerCache for this field
                    // SlicerCaches.Add(source, sourceField, name, slicerCacheType)
                    // source = PivotTable object
                    // sourceField = field name string
                    // name = cache name (optional, auto-generated if not provided)

                    // For regular PivotTables, use field name directly
                    // For OLAP, may need the hierarchical name
                    ct.ThrowIfCancellationRequested();
                    slicerCache = slicerCaches.Add2(pivot, fieldName);
                }

                // Get destination sheet and calculate position from cell reference
                sheets = ctx.Book.Worksheets;
                destSheet = sheets[destinationSheet];
                destRange = destSheet.Range[position];

                // Get position in points from the cell reference
                double top = Convert.ToDouble(destRange.Top);
                double left = Convert.ToDouble(destRange.Left);

                // Add visual Slicer to the cache
                // Slicers.Add(SlicerDestination, Level, Name, Caption, Top, Left, Width, Height)
                // For non-OLAP sources, Level should be Type.Missing or omitted
                slicers = slicerCache.Slicers;
                ct.ThrowIfCancellationRequested();
                slicer = slicers.Add(destSheet, Type.Missing, slicerName, slicerName, top, left);

                // Build result
                var result = BuildSlicerResult(slicer, slicerCache, fieldName, ct);
                result.Success = true;
                result.WorkflowHint = $"Slicer '{slicerName}' created for field '{fieldName}'. Use SetSlicerSelection to filter data, or connect additional PivotTables to this slicer.";

                return result;
            }
            finally
            {
                ComUtilities.Release(ref slicer);
                ComUtilities.Release(ref slicers);
                ComUtilities.Release(ref destRange);
                ComUtilities.Release(ref destSheet);
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref slicerCache);
                ComUtilities.Release(ref slicerCaches);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Lists all slicers in the workbook, optionally filtered by PivotTable
    /// </summary>
    public SlicerListResult ListSlicers(IExcelBatch batch, string? pivotTableName = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            var result = new SlicerListResult { Success = true };
            dynamic? slicerCaches = null;
            dynamic? targetPivot = null;

            try
            {
                slicerCaches = ctx.Book.SlicerCaches;

                // If filtering by PivotTable, find it first
                if (!string.IsNullOrEmpty(pivotTableName))
                {
                    targetPivot = FindPivotTable(ctx.Book, pivotTableName);
                }

                for (int cacheIndex = 1; cacheIndex <= slicerCaches.Count; cacheIndex++)
                {
                    ct.ThrowIfCancellationRequested();
                    dynamic? cache = null;
                    dynamic? slicers = null;

                    try
                    {
                        cache = slicerCaches.Item(cacheIndex);

                        // If filtering by PivotTable, check if this cache is connected
                        if (targetPivot != null && !IsSlicerCacheConnectedToPivot(cache, targetPivot, ct))
                        {
                            continue;
                        }

                        slicers = cache.Slicers;
                        for (int slicerIndex = 1; slicerIndex <= slicers.Count; slicerIndex++)
                        {
                            ct.ThrowIfCancellationRequested();
                            dynamic? slicer = null;
                            try
                            {
                                slicer = slicers.Item(slicerIndex);
                                var slicerInfo = BuildSlicerInfo(slicer, cache, ct);
                                result.Slicers.Add(slicerInfo);
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

                return result;
            }
            finally
            {
                ComUtilities.Release(ref targetPivot);
                ComUtilities.Release(ref slicerCaches);
            }
        });
    }

    /// <summary>
    /// Sets the selection for a slicer
    /// </summary>
    public SlicerResult SetSlicerSelection(IExcelBatch batch, string slicerName,
        List<string> selectedItems, bool clearFirst = true)
    {
        ArgumentNullException.ThrowIfNull(selectedItems);

        return batch.Execute((ctx, ct) =>
        {
            dynamic? slicerCaches = null;
            Excel.SlicerCache? targetCache = null;
            Excel.Slicer? targetSlicer = null;

            try
            {
                slicerCaches = ctx.Book.SlicerCaches;

                // Find the slicer by name
                var searchResult = FindSlicerByName(slicerCaches, slicerName, ct);
                targetCache = searchResult.Cache;
                targetSlicer = searchResult.Slicer;

                if (targetSlicer == null || targetCache == null)
                {
                    return new SlicerResult
                    {
                        Success = false,
                        ErrorMessage = $"Slicer '{slicerName}' not found in workbook"
                    };
                }

                // If no items specified, select all (clear filter)
                bool selectAll = selectedItems.Count == 0;

                if (targetCache.OLAP)
                {
                    SetOlapSlicerSelection(targetSlicer, targetCache,
                        selectedItems, clearFirst, ct);
                }
                else
                {
                    SlicerSelection.SetNonOlapSelection(targetCache, selectedItems, clearFirst, ct);
                }

                // Build result with updated state
                string fieldName = GetSlicerCacheFieldName(targetCache);
                var result = BuildSlicerResult(targetSlicer, targetCache, fieldName, ct);
                result.Success = true;
                result.WorkflowHint = selectAll
                    ? $"Slicer '{slicerName}' filter cleared - all items are now visible."
                    : $"Slicer '{slicerName}' selection updated to {result.SelectedItems.Count} item(s).";

                return result;
            }
            finally
            {
                ComUtilities.Release(ref targetSlicer);
                ComUtilities.Release(ref targetCache);
                ComUtilities.Release(ref slicerCaches);
            }
        });
    }

    /// <summary>
    /// Deletes a slicer from the workbook
    /// </summary>
    public OperationResult DeleteSlicer(IExcelBatch batch, string slicerName)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? slicerCaches = null;
            dynamic? targetCache = null;
            dynamic? targetSlicer = null;

            try
            {
                slicerCaches = ctx.Book.SlicerCaches;

                // Find the slicer by name
                var searchResult = FindSlicerByName(slicerCaches, slicerName, ct);
                targetCache = searchResult.Cache;
                targetSlicer = searchResult.Slicer;

                if (targetSlicer == null)
                {
                    return new OperationResult
                    {
                        Success = false,
                        ErrorMessage = $"Slicer '{slicerName}' not found in workbook"
                    };
                }

                // Delete the visual slicer
                targetSlicer.Delete();

                // Note: The SlicerCache will be automatically deleted if this was the last slicer
                // connected to it. Excel handles this automatically.

                return new OperationResult { Success = true };
            }
            finally
            {
                ComUtilities.Release(ref targetSlicer);
                ComUtilities.Release(ref targetCache);
                ComUtilities.Release(ref slicerCaches);
            }
        });
    }

    #region Slicer Helper Methods

    /// <summary>
    /// Result of searching for a slicer by name (avoids dynamic tuple deconstruction)
    /// </summary>
    private readonly struct SlicerSearchResult
    {
        public dynamic? Cache { get; init; }
        public dynamic? Slicer { get; init; }
    }

    /// <summary>
    /// Finds an existing SlicerCache for a field on a specific PivotTable
    /// </summary>
    private static dynamic? FindExistingSlicerCache(dynamic slicerCaches, dynamic pivot, string fieldName, CancellationToken ct)
    {
        for (int i = 1; i <= slicerCaches.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            dynamic? cache = null;
            bool found = false;
            try
            {
                cache = slicerCaches.Item(i);

                // Check if cache is for the same field
                string cacheFieldName = GetSlicerCacheFieldName(cache);
                if (!string.Equals(cacheFieldName, fieldName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                // Check if this cache is connected to our PivotTable
                if (IsSlicerCacheConnectedToPivot(cache, pivot, ct))
                {
                    found = true;
                    return cache; // Don't release - returning to caller
                }
            }
            finally
            {
                if (!found)
                    ComUtilities.Release(ref cache);
            }
        }

        return null;
    }

    /// <summary>
    /// Gets the source field name from a SlicerCache
    /// </summary>
    private static string GetSlicerCacheFieldName(dynamic cache)
    {
        return cache.SourceName;
    }

    /// <summary>
    /// Checks if a SlicerCache is connected to a specific PivotTable.
    /// Returns false for Table slicers (cache.List == true) since they don't connect to PivotTables.
    /// </summary>
    private static bool IsSlicerCacheConnectedToPivot(dynamic cache, dynamic targetPivot, CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();
        // Per MS docs: List property is true for Table slicers, false for PivotTable slicers
        // https://learn.microsoft.com/en-us/office/vba/api/excel.slicercache.list
        // Table slicers don't connect to PivotTables
        if (cache.List == true)
        {
            return false;
        }

        dynamic? pivotTables = null;
        try
        {
            pivotTables = cache.PivotTables;
            string targetName = targetPivot.Name?.ToString() ?? string.Empty;

            for (int i = 1; i <= pivotTables.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                dynamic? pt = null;
                try
                {
                    pt = pivotTables.Item(i);
                    string ptName = pt.Name?.ToString() ?? string.Empty;
                    ct.ThrowIfCancellationRequested();
                    if (string.Equals(ptName, targetName, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref pt);
                }
            }
            return false;
        }
        finally
        {
            ComUtilities.Release(ref pivotTables);
        }
    }

    /// <summary>
    /// Finds a slicer by name across all SlicerCaches
    /// </summary>
    private static SlicerSearchResult FindSlicerByName(dynamic slicerCaches, string slicerName, CancellationToken ct)
    {
        for (int cacheIndex = 1; cacheIndex <= slicerCaches.Count; cacheIndex++)
        {
            ct.ThrowIfCancellationRequested();
            dynamic? cache = null;
            dynamic? slicers = null;
            bool found = false;

            try
            {
                cache = slicerCaches.Item(cacheIndex);
                slicers = cache.Slicers;

                for (int slicerIndex = 1; slicerIndex <= slicers.Count; slicerIndex++)
                {
                    ct.ThrowIfCancellationRequested();
                    dynamic? slicer = null;
                    try
                    {
                        slicer = slicers.Item(slicerIndex);
                        string name = slicer.Name?.ToString() ?? string.Empty;

                        if (string.Equals(name, slicerName, StringComparison.OrdinalIgnoreCase))
                        {
                            // Found it - return both cache and slicer (don't release)
                            found = true;
                            return new SlicerSearchResult { Cache = cache, Slicer = slicer };
                        }
                    }
                    finally
                    {
                        if (!found)
                            ComUtilities.Release(ref slicer);
                    }
                }
            }
            finally
            {
                ComUtilities.Release(ref slicers);
                if (!found)
                    ComUtilities.Release(ref cache);
            }
        }

        return new SlicerSearchResult { Cache = null, Slicer = null };
    }

    /// <summary>
    /// Builds a SlicerInfo from COM objects
    /// </summary>
    private static SlicerInfo BuildSlicerInfo(dynamic slicer, dynamic cache, CancellationToken ct)
    {
        var info = new SlicerInfo
        {
            Name = slicer.Name?.ToString() ?? string.Empty,
            Caption = slicer.Caption?.ToString() ?? string.Empty,
            FieldName = GetSlicerCacheFieldName(cache),
            ColumnCount = Convert.ToInt32(slicer.NumberOfColumns)
        };

        // Get sheet name and position
        dynamic? parent = null;
        try
        {
            parent = slicer.Parent;
            info.SheetName = parent.Name?.ToString() ?? string.Empty;
        }
        finally
        {
            ComUtilities.Release(ref parent);
        }

        // Get position (top-left cell) - per Microsoft docs, TopLeftCell is on Shape object
        // https://learn.microsoft.com/en-us/office/vba/api/excel.shape.topleftcell
        dynamic? shape = null;
        dynamic? topLeftCell = null;
        try
        {
            shape = slicer.Shape;
            topLeftCell = shape.TopLeftCell;
            info.Position = topLeftCell?.Address?.ToString()?.Replace("$", "") ?? string.Empty;
        }
        finally
        {
            ComUtilities.Release(ref topLeftCell);
            ComUtilities.Release(ref shape);
        }

        // Get selected and available items from cache
        SlicerItemsResult items = GetSlicerItems((Excel.Slicer)slicer, (Excel.SlicerCache)cache, ct);
        info.SelectedItems = items.Selected;
        info.AvailableItems = items.Available;

        // Get connected PivotTables
        info.ConnectedPivotTables = GetConnectedPivotTableNames(cache, ct);

        return info;
    }

    /// <summary>
    /// Builds a SlicerResult from COM objects
    /// </summary>
    private static SlicerResult BuildSlicerResult(dynamic slicer, dynamic cache, string fieldName, CancellationToken ct)
    {
        var result = new SlicerResult
        {
            Name = slicer.Name?.ToString() ?? string.Empty,
            Caption = slicer.Caption?.ToString() ?? string.Empty,
            FieldName = fieldName
        };

        // Get sheet name and position
        dynamic? parent = null;
        try
        {
            parent = slicer.Parent;
            result.SheetName = parent.Name?.ToString() ?? string.Empty;
        }
        finally
        {
            ComUtilities.Release(ref parent);
        }

        // Get position - per Microsoft docs, TopLeftCell is on Shape object
        // https://learn.microsoft.com/en-us/office/vba/api/excel.shape.topleftcell
        dynamic? shape = null;
        dynamic? topLeftCell = null;
        try
        {
            shape = slicer.Shape;
            topLeftCell = shape.TopLeftCell;
            result.Position = topLeftCell?.Address?.ToString()?.Replace("$", "") ?? string.Empty;
        }
        finally
        {
            ComUtilities.Release(ref topLeftCell);
            ComUtilities.Release(ref shape);
        }

        // Get items
        SlicerItemsResult items = GetSlicerItems((Excel.Slicer)slicer, (Excel.SlicerCache)cache, ct);
        result.SelectedItems = items.Selected;
        result.AvailableItems = items.Available;

        // Get connected PivotTables
        result.ConnectedPivotTables = GetConnectedPivotTableNames(cache, ct);

        return result;
    }

    /// <summary>
    /// Result of getting slicer items (avoids dynamic tuple deconstruction)
    /// </summary>
    private readonly struct SlicerItemsResult
    {
        public List<string> Selected { get; init; }
        public List<string> Available { get; init; }
    }

    /// <summary>
    /// Gets selected and available items from a SlicerCache
    /// </summary>
    private static SlicerItemsResult GetSlicerItems(Excel.Slicer slicer, Excel.SlicerCache cache, CancellationToken ct)
    {
        var items = ReadSlicerItems(slicer, cache, ct);
        return new SlicerItemsResult
        {
            Selected = items.Where(item => item.Selected).Select(item => item.Caption).ToList(),
            Available = items.Select(item => item.Caption).ToList()
        };
    }

    internal sealed record SlicerItemState(string Name, string Caption, bool Selected);

    internal static string ResolveOlapSlicerItemName(IReadOnlyList<SlicerItemState> items, string requested)
    {
        var matches = items.Where(item => string.Equals(item.Name, requested, StringComparison.OrdinalIgnoreCase)).ToList();
        if (matches.Count == 0)
            matches = items.Where(item => string.Equals(item.Caption, requested, StringComparison.OrdinalIgnoreCase)).ToList();
        if (matches.Count == 0)
            throw new ArgumentException($"Slicer item '{requested}' was not found. Use availableItems from list-slicers.", nameof(requested));
        if (matches.Count > 1)
        {
            string candidates = JsonSerializer.Serialize(matches.Select(item => item.Name));
            throw new ArgumentException(
                $"Slicer caption '{requested}' is ambiguous. Retry with one of these matching MDX unique names: {candidates}",
                nameof(requested));
        }
        return matches[0].Name;
    }

    private static List<SlicerItemState> ReadSlicerItems(Excel.Slicer slicer, Excel.SlicerCache cache,
        CancellationToken ct)
    {
        var result = new List<SlicerItemState>();
        Excel.SlicerCacheLevel? level = null;
        Excel.SlicerItems? items = null;
        try
        {
            bool olap = cache.OLAP;
            if (olap)
            {
                // OLAP items belong to the visual slicer's level, not SlicerCache.SlicerItems.
                level = slicer.SlicerCacheLevel;
                items = level.SlicerItems;
            }
            else
            {
                items = cache.SlicerItems;
            }

            for (int i = 1; i <= items.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.SlicerItem? item = null;
                try
                {
                    item = items.Item[i];
                    result.Add(new SlicerItemState(item.Name, olap ? item.Caption : item.Name, item.Selected));
                }
                finally
                {
                    ComUtilities.Release(ref item);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref items);
            ComUtilities.Release(ref level);
        }

        return result;
    }

    private static void SetOlapSlicerSelection(Excel.Slicer slicer, Excel.SlicerCache cache,
        List<string> selectedItems, bool clearFirst, CancellationToken ct)
    {
        if (selectedItems.Count == 0)
        {
            ct.ThrowIfCancellationRequested();
            cache.ClearManualFilter();
            return;
        }

        var items = ReadSlicerItems(slicer, cache, ct);
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string requested in selectedItems)
        {
            ct.ThrowIfCancellationRequested();
            names.Add(ResolveOlapSlicerItemName(items, requested));
        }

        if (!clearFirst)
        {
            names.UnionWith(items.Where(item => item.Selected).Select(item => item.Name));
            // Preserve manual filters on other levels of a shared OLAP hierarchy.
            if (!cache.FilterCleared)
            {
                if (cache.VisibleSlicerItemsList is not Array visible)
                    throw new InvalidOperationException("Excel did not return the OLAP slicer's manual selection.");
                names.UnionWith(visible.Cast<string>());
            }
        }
        ct.ThrowIfCancellationRequested();
        cache.VisibleSlicerItemsList = names.ToArray();
    }

    /// <summary>
    /// Gets names of PivotTables connected to a SlicerCache.
    /// Returns empty list for Table slicers (cache.List == true).
    /// </summary>
    private static List<string> GetConnectedPivotTableNames(dynamic cache, CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();
        var names = new List<string>();

        // Per MS docs: List property is true for Table slicers
        // Table slicers don't have PivotTables collection
        if (cache.List == true)
        {
            return names;
        }

        dynamic? pivotTables = null;
        try
        {
            pivotTables = cache.PivotTables;

            for (int i = 1; i <= pivotTables.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                dynamic? pt = null;
                try
                {
                    pt = pivotTables.Item(i);
                    string name = pt.Name?.ToString() ?? string.Empty;
                    ct.ThrowIfCancellationRequested();
                    if (!string.IsNullOrEmpty(name))
                    {
                        names.Add(name);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref pt);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref pivotTables);
        }

        return names;
    }

    #endregion
}
