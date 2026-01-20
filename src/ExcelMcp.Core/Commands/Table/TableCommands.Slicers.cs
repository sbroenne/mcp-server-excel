using System.Runtime.InteropServices;

using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Table slicer operations (CreateTableSlicer, ListTableSlicers, SetTableSlicerSelection, DeleteTableSlicer)
/// </summary>
public partial class TableCommands
{

    /// <inheritdoc />
    public SlicerResult CreateTableSlicer(IExcelBatch batch, string tableName,
        string columnName, string slicerName, string destinationSheet, string position)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? slicerCaches = null;
            dynamic? slicerCache = null;
            dynamic? slicers = null;
            dynamic? slicer = null;
            dynamic? destSheet = null;
            dynamic? destRange = null;

            try
            {
                table = FindTable(ctx.Book, tableName);
                slicerCaches = ctx.Book.SlicerCaches;

                // Validate the column exists in the table before creating slicer
                if (!TableColumnExists(table, columnName))
                {
                    return new SlicerResult
                    {
                        Success = false,
                        ErrorMessage = $"Column '{columnName}' not found in table '{tableName}'"
                    };
                }

                // Check if a SlicerCache already exists for this column on this table
                slicerCache = FindExistingTableSlicerCache(slicerCaches, table, columnName);

                if (slicerCache == null)
                {
                    // Create new SlicerCache for this column
                    // Use the deprecated SlicerCaches.Add(source, sourceField) method for Table slicers
                    // The Add method (without SlicerCacheType) accepts ListObject as source
                    // Note: Add2 does NOT accept ListObject per Microsoft documentation
                    slicerCache = slicerCaches.Add(table, columnName);
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
                var result = BuildTableSlicerResult(slicer, slicerCache, columnName, tableName);
                result.Success = true;
                result.WorkflowHint = $"Slicer '{slicerName}' created for column '{columnName}' in table '{tableName}'. Use SetTableSlicerSelection to filter data.";

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
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
    public SlicerListResult ListTableSlicers(IExcelBatch batch, string? tableName = null)
    {
        // Security: Validate table name if provided
        if (!string.IsNullOrEmpty(tableName))
        {
            ValidateTableName(tableName);
        }

        return batch.Execute((ctx, ct) =>
        {
            var result = new SlicerListResult { Success = true };
            dynamic? slicerCaches = null;
            dynamic? targetTable = null;

            try
            {
                slicerCaches = ctx.Book.SlicerCaches;

                // If filtering by Table, find it first
                if (!string.IsNullOrEmpty(tableName))
                {
                    targetTable = FindTable(ctx.Book, tableName);
                }

                for (int cacheIndex = 1; cacheIndex <= slicerCaches.Count; cacheIndex++)
                {
                    dynamic? cache = null;
                    dynamic? slicers = null;

                    try
                    {
                        cache = slicerCaches.Item(cacheIndex);

                        // Check if this is a Table slicer using the List boolean property
                        // SlicerCache.List returns true if the slicer is connected to a ListObject (Table)
                        bool isTableSlicer = false;
                        try
                        {
                            isTableSlicer = cache.List == true;
                        }
                        catch (COMException)
                        {
                            // List property not available - not a Table slicer
                            isTableSlicer = false;
                        }

                        if (!isTableSlicer)
                        {
                            continue; // Skip non-Table slicers (e.g., PivotTable slicers)
                        }

                        // If filtering by Table, check if this cache is connected to it
                        if (targetTable != null && !IsSlicerCacheConnectedToTable(cache, targetTable))
                        {
                            continue;
                        }

                        // Get the connected table name
                        string connectedTableName = GetSlicerCacheTableName(cache);

                        slicers = cache.Slicers;
                        for (int slicerIndex = 1; slicerIndex <= slicers.Count; slicerIndex++)
                        {
                            dynamic? slicer = null;
                            try
                            {
                                slicer = slicers.Item(slicerIndex);
                                var slicerInfo = BuildTableSlicerInfo(slicer, cache, connectedTableName);
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
                ComUtilities.Release(ref targetTable);
                ComUtilities.Release(ref slicerCaches);
            }
        });
    }

    /// <inheritdoc />
    public SlicerResult SetTableSlicerSelection(IExcelBatch batch, string slicerName,
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

                // Find the slicer by name (searching Table slicers only)
                var searchResult = FindTableSlicerByName(slicerCaches, slicerName);
                targetCache = searchResult.Cache;
                targetSlicer = searchResult.Slicer;

                if (targetSlicer == null || targetCache == null)
                {
                    return new SlicerResult
                    {
                        Success = false,
                        ErrorMessage = $"Table slicer '{slicerName}' not found in workbook"
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
                string columnName = GetSlicerCacheColumnName(targetCache);
                string connectedTableName = GetSlicerCacheTableName(targetCache);
                var result = BuildTableSlicerResult(targetSlicer, targetCache, columnName, connectedTableName);
                result.Success = true;
                result.WorkflowHint = selectAll
                    ? $"Table slicer '{slicerName}' filter cleared - all items are now visible."
                    : $"Table slicer '{slicerName}' selection updated to {selectedItems.Count} item(s).";

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

    /// <inheritdoc />
    public OperationResult DeleteTableSlicer(IExcelBatch batch, string slicerName)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? slicerCaches = null;
            dynamic? targetCache = null;
            dynamic? targetSlicer = null;

            try
            {
                slicerCaches = ctx.Book.SlicerCaches;

                // Find the slicer by name (searching Table slicers only)
                var searchResult = FindTableSlicerByName(slicerCaches, slicerName);
                targetCache = searchResult.Cache;
                targetSlicer = searchResult.Slicer;

                if (targetSlicer == null)
                {
                    return new OperationResult
                    {
                        Success = false,
                        ErrorMessage = $"Table slicer '{slicerName}' not found in workbook"
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

    #region Table Slicer Helper Methods

    /// <summary>
    /// Checks if a column with the given name exists in the table
    /// </summary>
    /// <param name="table">The table (ListObject) to check</param>
    /// <param name="columnName">The column name to search for</param>
    /// <returns>True if the column exists, false otherwise</returns>
    private static bool TableColumnExists(dynamic table, string columnName)
    {
        dynamic? listColumns = null;
        try
        {
            listColumns = table.ListColumns;
            for (int i = 1; i <= listColumns.Count; i++)
            {
                dynamic? column = null;
                try
                {
                    column = listColumns.Item(i);
                    string name = column.Name?.ToString() ?? string.Empty;
                    if (string.Equals(name, columnName, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref column);
                }
            }
            return false;
        }
        finally
        {
            ComUtilities.Release(ref listColumns);
        }
    }

    /// <summary>
    /// Result of searching for a table slicer by name
    /// </summary>
    private readonly struct TableSlicerSearchResult
    {
        public dynamic? Cache { get; init; }
        public dynamic? Slicer { get; init; }
    }

    /// <summary>
    /// Finds an existing SlicerCache for a column on a specific Table
    /// </summary>
    private static dynamic? FindExistingTableSlicerCache(dynamic slicerCaches, dynamic table, string columnName)
    {
        for (int i = 1; i <= slicerCaches.Count; i++)
        {
            dynamic? cache = null;
            try
            {
                cache = slicerCaches.Item(i);

                // Check if this is a Table slicer using the List boolean property
                bool isTableSlicer = false;
                try
                {
                    isTableSlicer = cache.List == true;
                }
                catch (COMException)
                {
                    // List property doesn't exist on PivotTable slicer caches
                    isTableSlicer = false;
                }

                if (!isTableSlicer)
                {
                    ComUtilities.Release(ref cache);
                    continue;
                }

                // Check if cache is for the same column
                string cacheColumnName = GetSlicerCacheColumnName(cache);
                if (!string.Equals(cacheColumnName, columnName, StringComparison.OrdinalIgnoreCase))
                {
                    ComUtilities.Release(ref cache);
                    continue;
                }

                // Check if this cache is connected to our Table
                if (IsSlicerCacheConnectedToTable(cache, table))
                {
                    return cache; // Don't release - returning to caller
                }

                ComUtilities.Release(ref cache);
            }
            catch (COMException)
            {
                // COM access may fail for certain cache types - continue searching
                ComUtilities.Release(ref cache);
            }
        }

        return null;
    }

    /// <summary>
    /// Gets the source column name from a Table SlicerCache
    /// </summary>
    private static string GetSlicerCacheColumnName(dynamic cache)
    {
        try
        {
            // For Table slicers, SourceName contains the column name
            string? sourceName = cache.SourceName?.ToString();
            if (!string.IsNullOrEmpty(sourceName))
                return sourceName;

            // Fallback: parse from cache name (usually "Slicer_ColumnName" format)
            string cacheName = cache.Name?.ToString() ?? string.Empty;
            if (cacheName.StartsWith("Slicer_", StringComparison.OrdinalIgnoreCase))
            {
                return cacheName[7..]; // Remove "Slicer_" prefix
            }
            return cacheName;
        }
        catch (COMException)
        {
            // COM access may fail for certain cache configurations
            return "Unknown";
        }
    }

    /// <summary>
    /// Gets the connected Table name from a SlicerCache
    /// </summary>
    private static string GetSlicerCacheTableName(dynamic cache)
    {
        dynamic? listObject = null;
        try
        {
            // For Table slicers, the cache has a ListObject property
            listObject = cache.ListObject;
            if (listObject != null)
            {
                return listObject.Name?.ToString() ?? "Unknown";
            }
            return "Unknown";
        }
        catch (COMException)
        {
            // COM access may fail for certain cache configurations
            return "Unknown";
        }
        finally
        {
            ComUtilities.Release(ref listObject);
        }
    }

    /// <summary>
    /// Checks if a SlicerCache is connected to a specific Table
    /// </summary>
    private static bool IsSlicerCacheConnectedToTable(dynamic cache, dynamic table)
    {
        dynamic? cacheListObject = null;
        try
        {
            cacheListObject = cache.ListObject;
            if (cacheListObject == null)
                return false;

            string cacheTableName = cacheListObject.Name?.ToString() ?? string.Empty;
            string targetTableName = table.Name?.ToString() ?? string.Empty;

            return string.Equals(cacheTableName, targetTableName, StringComparison.OrdinalIgnoreCase);
        }
        catch (COMException)
        {
            // COM access may fail for certain cache configurations
            return false;
        }
        finally
        {
            ComUtilities.Release(ref cacheListObject);
        }
    }

    /// <summary>
    /// Finds a Table slicer by name
    /// </summary>
    private static TableSlicerSearchResult FindTableSlicerByName(dynamic slicerCaches, string slicerName)
    {
        for (int cacheIndex = 1; cacheIndex <= slicerCaches.Count; cacheIndex++)
        {
            dynamic? cache = null;
            dynamic? slicers = null;

            try
            {
                cache = slicerCaches.Item(cacheIndex);

                // Check if this is a Table slicer using the List boolean property
                bool isTableSlicer = false;
                try
                {
                    isTableSlicer = cache.List == true;
                }
                catch (COMException)
                {
                    // List property doesn't exist on PivotTable slicer caches
                    isTableSlicer = false;
                }

                if (!isTableSlicer)
                {
                    ComUtilities.Release(ref cache);
                    continue;
                }

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
                            // Don't release cache and slicer - returning to caller
                            ComUtilities.Release(ref slicers);
                            return new TableSlicerSearchResult { Cache = cache, Slicer = slicer };
                        }

                        ComUtilities.Release(ref slicer);
                    }
                    catch (COMException)
                    {
                        // COM access failed for this slicer, continue searching
                        ComUtilities.Release(ref slicer);
                    }
                }

                ComUtilities.Release(ref slicers);
                ComUtilities.Release(ref cache);
            }
            catch (COMException)
            {
                // COM access failed for this cache, continue searching
                ComUtilities.Release(ref slicers);
                ComUtilities.Release(ref cache);
            }
        }

        return new TableSlicerSearchResult { Cache = null, Slicer = null };
    }

    /// <summary>
    /// Builds a SlicerInfo object for a Table slicer
    /// </summary>
    private static SlicerInfo BuildTableSlicerInfo(dynamic slicer, dynamic cache, string tableName)
    {
        dynamic? parent = null;
        dynamic? slicerItems = null;

        try
        {
            parent = slicer.Parent;
            slicerItems = cache.SlicerItems;

            var info = new SlicerInfo
            {
                Name = slicer.Name?.ToString() ?? string.Empty,
                Caption = slicer.Caption?.ToString() ?? string.Empty,
                FieldName = GetSlicerCacheColumnName(cache),
                SheetName = parent?.Name?.ToString() ?? string.Empty,
                Position = GetSlicerPosition(slicer),
                ColumnCount = Convert.ToInt32(slicer.NumberOfColumns ?? 1),
                ConnectedTable = tableName,
                SourceType = "Table"
            };

            // Collect selected and available items
            for (int i = 1; i <= slicerItems.Count; i++)
            {
                dynamic? item = null;
                try
                {
                    item = slicerItems.Item(i);
                    string itemName = item.Name?.ToString() ?? string.Empty;
                    info.AvailableItems.Add(itemName);

                    bool isSelected = item.Selected;
                    if (isSelected)
                    {
                        info.SelectedItems.Add(itemName);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref item);
                }
            }

            return info;
        }
        finally
        {
            ComUtilities.Release(ref slicerItems);
            ComUtilities.Release(ref parent);
        }
    }

    /// <summary>
    /// Builds a SlicerResult object for a Table slicer operation
    /// </summary>
    private static SlicerResult BuildTableSlicerResult(dynamic slicer, dynamic cache, string columnName, string tableName)
    {
        dynamic? parent = null;
        dynamic? slicerItems = null;

        try
        {
            parent = slicer.Parent;
            slicerItems = cache.SlicerItems;

            var result = new SlicerResult
            {
                Name = slicer.Name?.ToString() ?? string.Empty,
                Caption = slicer.Caption?.ToString() ?? string.Empty,
                FieldName = columnName,
                SheetName = parent?.Name?.ToString() ?? string.Empty,
                Position = GetSlicerPosition(slicer),
                ConnectedTable = tableName,
                SourceType = "Table"
            };

            // Collect selected and available items
            for (int i = 1; i <= slicerItems.Count; i++)
            {
                dynamic? item = null;
                try
                {
                    item = slicerItems.Item(i);
                    string itemName = item.Name?.ToString() ?? string.Empty;
                    result.AvailableItems.Add(itemName);

                    bool isSelected = item.Selected;
                    if (isSelected)
                    {
                        result.SelectedItems.Add(itemName);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref item);
                }
            }

            return result;
        }
        finally
        {
            ComUtilities.Release(ref slicerItems);
            ComUtilities.Release(ref parent);
        }
    }

    /// <summary>
    /// Gets the position string for a slicer (e.g., "H2")
    /// </summary>
    private static string GetSlicerPosition(dynamic slicer)
    {
        dynamic? topLeftCell = null;
        try
        {
            topLeftCell = slicer.TopLeftCell;
            if (topLeftCell != null)
            {
                return topLeftCell.Address?.ToString()?.Replace("$", "") ?? string.Empty;
            }
            return string.Empty;
        }
        catch (COMException)
        {
            // COM access may fail for certain slicer configurations
            return string.Empty;
        }
        finally
        {
            ComUtilities.Release(ref topLeftCell);
        }
    }

    #endregion
}
