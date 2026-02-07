using System.Runtime.InteropServices;

using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

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

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                slicerCaches = ctx.Book.SlicerCaches;

                // Check if a SlicerCache already exists for this field+PivotTable
                // If so, we add a new visual Slicer to the existing cache
                slicerCache = FindExistingSlicerCache(slicerCaches, pivot, fieldName);

                if (slicerCache == null)
                {
                    // Create new SlicerCache for this field
                    // SlicerCaches.Add(source, sourceField, name, slicerCacheType)
                    // source = PivotTable object
                    // sourceField = field name string
                    // name = cache name (optional, auto-generated if not provided)

                    // For regular PivotTables, use field name directly
                    // For OLAP, may need the hierarchical name
                    slicerCache = slicerCaches.Add2(pivot, fieldName);
                }

                // Get destination sheet and calculate position from cell reference
                destSheet = ctx.Book.Worksheets.Item(destinationSheet);
                destRange = destSheet.Range[position];

                // Get position in points from the cell reference
                double top = Convert.ToDouble(destRange.Top);
                double left = Convert.ToDouble(destRange.Left);

                // Add visual Slicer to the cache
                // Slicers.Add(SlicerDestination, Level, Name, Caption, Top, Left, Width, Height)
                // For non-OLAP sources, Level should be Type.Missing or omitted
                slicers = slicerCache.Slicers;
                slicer = slicers.Add(destSheet, Type.Missing, slicerName, slicerName, top, left);

                // Build result
                var result = BuildSlicerResult(slicer, slicerCache, fieldName);
                result.Success = true;
                result.WorkflowHint = $"Slicer '{slicerName}' created for field '{fieldName}'. Use SetSlicerSelection to filter data, or connect additional PivotTables to this slicer.";

                return result;
            }
            finally
            {
                ComUtilities.Release(ref destRange);
                ComUtilities.Release(ref destSheet);
                ComUtilities.Release(ref slicer);
                ComUtilities.Release(ref slicers);
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
                    dynamic? cache = null;
                    dynamic? slicers = null;

                    try
                    {
                        cache = slicerCaches.Item(cacheIndex);

                        // If filtering by PivotTable, check if this cache is connected
                        if (targetPivot != null && !IsSlicerCacheConnectedToPivot(cache, targetPivot))
                        {
                            continue;
                        }

                        slicers = cache.Slicers;
                        for (int slicerIndex = 1; slicerIndex <= slicers.Count; slicerIndex++)
                        {
                            dynamic? slicer = null;
                            try
                            {
                                slicer = slicers.Item(slicerIndex);
                                var slicerInfo = BuildSlicerInfo(slicer, cache);
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
        return batch.Execute((ctx, ct) =>
        {
            dynamic? slicerCaches = null;
            dynamic? targetCache = null;
            dynamic? targetSlicer = null;
            dynamic? slicerItems = null;

            try
            {
                slicerCaches = ctx.Book.SlicerCaches;

                // Find the slicer by name
                var searchResult = FindSlicerByName(slicerCaches, slicerName);
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

                // Get slicer items from the cache
                slicerItems = targetCache.SlicerItems;

                // Build set of items to select for fast lookup
                var itemsToSelect = new HashSet<string>(selectedItems, StringComparer.OrdinalIgnoreCase);

                // If no items specified, select all (clear filter)
                bool selectAll = selectedItems.Count == 0;

                // Iterate through slicer items and set selection
                for (int i = 1; i <= slicerItems.Count; i++)
                {
                    dynamic? item = null;
                    try
                    {
                        item = slicerItems.Item(i);
                        string itemName = item.Name?.ToString() ?? string.Empty;

                        if (selectAll)
                        {
                            item.Selected = true;
                        }
                        else if (clearFirst)
                        {
                            // Clear first mode: select only specified items
                            item.Selected = itemsToSelect.Contains(itemName);
                        }
                        else
                        {
                            // Additive mode: add to existing selection
                            if (itemsToSelect.Contains(itemName))
                            {
                                item.Selected = true;
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref item);
                    }
                }

                // Build result with updated state
                string fieldName = GetSlicerCacheFieldName(targetCache);
                var result = BuildSlicerResult(targetSlicer, targetCache, fieldName);
                result.Success = true;
                result.WorkflowHint = selectAll
                    ? $"Slicer '{slicerName}' filter cleared - all items are now visible."
                    : $"Slicer '{slicerName}' selection updated to {selectedItems.Count} item(s).";

                return result;
            }
            finally
            {
                ComUtilities.Release(ref slicerItems);
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
                var searchResult = FindSlicerByName(slicerCaches, slicerName);
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
    private static dynamic? FindExistingSlicerCache(dynamic slicerCaches, dynamic pivot, string fieldName)
    {
        for (int i = 1; i <= slicerCaches.Count; i++)
        {
            dynamic? cache = null;
            try
            {
                cache = slicerCaches.Item(i);

                // Check if cache is for the same field
                string cacheFieldName = GetSlicerCacheFieldName(cache);
                if (!string.Equals(cacheFieldName, fieldName, StringComparison.OrdinalIgnoreCase))
                {
                    ComUtilities.Release(ref cache);
                    continue;
                }

                // Check if this cache is connected to our PivotTable
                if (IsSlicerCacheConnectedToPivot(cache, pivot))
                {
                    return cache; // Don't release - returning to caller
                }

                ComUtilities.Release(ref cache);
            }
            catch (COMException)
            {
                // COM property access may fail for certain cache types - continue searching
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
        dynamic? sourceField = null;
        try
        {
            // Try to get SourceName first (OLAP), then fall back to checking PivotField
            try
            {
                string? sourceName = cache.SourceName?.ToString();
                if (!string.IsNullOrEmpty(sourceName))
                    return sourceName;
            }
            catch (COMException)
            {
                // SourceName property not available for this cache type - fall back to Name
            }

            // For regular slicers, get from PivotTables collection
            dynamic? pivotTables = null;
            try
            {
                pivotTables = cache.PivotTables;
                if (pivotTables != null && pivotTables.Count > 0)
                {
                    dynamic? pt = null;
                    try
                    {
                        pt = pivotTables.Item(1);
                        // The cache Name often contains the field name
                        string cacheName = cache.Name?.ToString() ?? string.Empty;
                        // SlicerCache names are typically "Slicer_FieldName" format
                        if (cacheName.StartsWith("Slicer_", StringComparison.OrdinalIgnoreCase))
                        {
                            return cacheName[7..]; // Remove "Slicer_" prefix
                        }
                        return cacheName;
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

            return cache.Name?.ToString() ?? "Unknown";
        }
        finally
        {
            ComUtilities.Release(ref sourceField);
        }
    }

    /// <summary>
    /// Checks if a SlicerCache is connected to a specific PivotTable.
    /// Returns false for Table slicers (cache.List == true) since they don't connect to PivotTables.
    /// </summary>
    private static bool IsSlicerCacheConnectedToPivot(dynamic cache, dynamic targetPivot)
    {
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
                dynamic? pt = null;
                try
                {
                    pt = pivotTables.Item(i);
                    string ptName = pt.Name?.ToString() ?? string.Empty;
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
    private static SlicerSearchResult FindSlicerByName(dynamic slicerCaches, string slicerName)
    {
        for (int cacheIndex = 1; cacheIndex <= slicerCaches.Count; cacheIndex++)
        {
            dynamic? cache = null;
            dynamic? slicers = null;

            try
            {
                cache = slicerCaches.Item(cacheIndex);
                slicers = cache.Slicers;

                for (int slicerIndex = 1; slicerIndex <= slicers.Count; slicerIndex++)
                {
                    dynamic? slicer = null;
                    try
                    {
                        slicer = slicers.Item(slicerIndex);
                        string name = slicer.Name?.ToString() ?? string.Empty;

                        if (string.Equals(name, slicerName, StringComparison.OrdinalIgnoreCase))
                        {
                            // Found it - return both cache and slicer (don't release)
                            ComUtilities.Release(ref slicers);
                            return new SlicerSearchResult { Cache = cache, Slicer = slicer };
                        }
                        ComUtilities.Release(ref slicer);
                    }
                    catch (COMException)
                    {
                        // COM access may fail for certain slicer types - continue searching
                        ComUtilities.Release(ref slicer);
                    }
                }

                ComUtilities.Release(ref slicers);
                ComUtilities.Release(ref cache);
            }
            catch (COMException)
            {
                // COM access may fail for certain cache types - continue searching
                ComUtilities.Release(ref slicers);
                ComUtilities.Release(ref cache);
            }
        }

        return new SlicerSearchResult { Cache = null, Slicer = null };
    }

    /// <summary>
    /// Builds a SlicerInfo from COM objects
    /// </summary>
    private static SlicerInfo BuildSlicerInfo(dynamic slicer, dynamic cache)
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
        var items = GetSlicerItems(cache);
        info.SelectedItems = items.Selected;
        info.AvailableItems = items.Available;

        // Get connected PivotTables
        info.ConnectedPivotTables = GetConnectedPivotTableNames(cache);

        return info;
    }

    /// <summary>
    /// Builds a SlicerResult from COM objects
    /// </summary>
    private static SlicerResult BuildSlicerResult(dynamic slicer, dynamic cache, string fieldName)
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
        var items = GetSlicerItems(cache);
        result.SelectedItems = items.Selected;
        result.AvailableItems = items.Available;

        // Get connected PivotTables
        result.ConnectedPivotTables = GetConnectedPivotTableNames(cache);

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
    private static SlicerItemsResult GetSlicerItems(dynamic cache)
    {
        var selected = new List<string>();
        var available = new List<string>();
        dynamic? slicerItems = null;

        try
        {
            slicerItems = cache.SlicerItems;

            for (int i = 1; i <= slicerItems.Count; i++)
            {
                dynamic? item = null;
                try
                {
                    item = slicerItems.Item(i);
                    string itemName = item.Name?.ToString() ?? string.Empty;

                    if (!string.IsNullOrEmpty(itemName))
                    {
                        available.Add(itemName);
                        if (item.Selected)
                        {
                            selected.Add(itemName);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref item);
                }
            }
        }
        catch (COMException)
        {
            // SlicerItems collection may not be accessible for certain cache types
        }
        finally
        {
            ComUtilities.Release(ref slicerItems);
        }

        return new SlicerItemsResult { Selected = selected, Available = available };
    }

    /// <summary>
    /// Gets names of PivotTables connected to a SlicerCache.
    /// Returns empty list for Table slicers (cache.List == true).
    /// </summary>
    private static List<string> GetConnectedPivotTableNames(dynamic cache)
    {
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
                dynamic? pt = null;
                try
                {
                    pt = pivotTables.Item(i);
                    string name = pt.Name?.ToString() ?? string.Empty;
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


