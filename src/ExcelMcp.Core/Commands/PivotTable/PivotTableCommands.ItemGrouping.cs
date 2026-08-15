using Excel = Microsoft.Office.Interop.Excel;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// Manual PivotTable item grouping and ungrouping.
/// </summary>
public partial class PivotTableCommands
{
    /// <inheritdoc />
    public PivotItemGroupResult GroupItems(
        IExcelBatch batch,
        string pivotTableName,
        string fieldName,
        List<string> itemNames,
        string groupName)
    {
        ArgumentNullException.ThrowIfNull(itemNames);
        if (itemNames.Count < 2)
        {
            throw new ArgumentException("At least two item names are required for manual grouping.", nameof(itemNames));
        }

        if (itemNames.Any(string.IsNullOrWhiteSpace))
        {
            throw new ArgumentException("Item names cannot be empty.", nameof(itemNames));
        }

        if (string.IsNullOrWhiteSpace(groupName))
        {
            throw new ArgumentException("Group name cannot be empty.", nameof(groupName));
        }

        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;
            Excel.PivotCache? cache = null;
            Excel.PivotField? field = null;
            Excel.Range? groupRange = null;
            object? groupOperationResult = null;
            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                cache = pivot.PivotCache();
                if (cache.OLAP)
                {
                    throw new InvalidOperationException(
                        "Manual item grouping is not supported for OLAP/Data Model PivotTables. " +
                        "Create the grouping column in the Data Model instead.");
                }

                field = (Excel.PivotField)pivot.PivotFields(fieldName);
                var orientation = field.Orientation;
                if (orientation is not Excel.XlPivotFieldOrientation.xlRowField
                    and not Excel.XlPivotFieldOrientation.xlColumnField)
                {
                    throw new InvalidOperationException(
                        $"Field '{fieldName}' must be placed in the Row or Column area before grouping items.");
                }

                var sourceItemNames = GetPivotItemNames(field);
                var fieldsBefore = GetPivotFieldItemSnapshot(pivot);
                foreach (var itemName in itemNames.Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.PivotItem? item = null;
                    Excel.Range? labelRange = null;
                    Excel.Range? unionRange = null;
                    try
                    {
                        item = (Excel.PivotItem)field.PivotItems(itemName);
                        labelRange = item.LabelRange;
                        if (groupRange == null)
                        {
                            groupRange = labelRange;
                            labelRange = null;
                        }
                        else
                        {
                            unionRange = ctx.App.Union(groupRange, labelRange);
                            ComUtilities.Release(ref groupRange);
                            groupRange = unionRange;
                            unionRange = null;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref unionRange);
                        ComUtilities.Release(ref labelRange);
                        ComUtilities.Release(ref item);
                    }
                }

                groupOperationResult = groupRange.Group();
                var fieldsAfter = GetPivotFieldItemSnapshot(pivot);
                var groupedFieldName = fieldsAfter
                    .Where(pair =>
                        !pair.Key.Equals(fieldName, StringComparison.OrdinalIgnoreCase)
                        && (!fieldsBefore.TryGetValue(pair.Key, out var previousItems)
                            || !previousItems.SetEquals(pair.Value)))
                    .Select(pair => pair.Key)
                    .Single();
                var baselineItems = fieldsBefore.TryGetValue(groupedFieldName, out var existingGroupedItems)
                    ? existingGroupedItems
                    : sourceItemNames;
                var generatedGroupItemName = fieldsAfter[groupedFieldName]
                    .Except(baselineItems, StringComparer.OrdinalIgnoreCase)
                    .Single();
                RenameGeneratedGroup(pivot, groupedFieldName, generatedGroupItemName, groupName);

                return new PivotItemGroupResult
                {
                    Success = true,
                    FieldName = fieldName,
                    GroupedFieldName = groupedFieldName,
                    GroupName = groupName,
                    Items = [.. itemNames],
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref groupOperationResult);
                ComUtilities.Release(ref groupRange);
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult UngroupField(
        IExcelBatch batch,
        string pivotTableName,
        string groupedFieldName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;
            Excel.PivotCache? cache = null;
            Excel.PivotField? field = null;
            Excel.Range? dataRange = null;
            object? ungroupResult = null;
            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                cache = pivot.PivotCache();
                if (cache.OLAP)
                {
                    throw new InvalidOperationException(
                        "Manual ungrouping is not supported for OLAP/Data Model PivotTables.");
                }

                field = (Excel.PivotField)pivot.PivotFields(groupedFieldName);
                dataRange = field.DataRange;
                ungroupResult = dataRange.Ungroup();
                if (GetPivotFieldItemSnapshot(pivot).ContainsKey(groupedFieldName))
                {
                    throw new InvalidOperationException(
                        $"Excel did not remove grouped field '{groupedFieldName}'.");
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "ungroup-field",
                    Message = $"Removed manual grouping from '{groupedFieldName}'.",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref ungroupResult);
                ComUtilities.Release(ref dataRange);
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    private static Dictionary<string, HashSet<string>> GetPivotFieldItemSnapshot(Excel.PivotTable pivot)
    {
        Excel.PivotFields? fields = null;
        try
        {
            fields = (Excel.PivotFields)pivot.PivotFields();
            var snapshot = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            for (var index = 1; index <= fields.Count; index++)
            {
                Excel.PivotField? field = null;
                try
                {
                    field = fields.Item(index);
                    var orientation = field.Orientation;
                    if (orientation is Excel.XlPivotFieldOrientation.xlRowField
                        or Excel.XlPivotFieldOrientation.xlColumnField)
                    {
                        snapshot[field.Name] = GetPivotItemNames(field);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref field);
                }
            }

            return snapshot;
        }
        finally
        {
            ComUtilities.Release(ref fields);
        }
    }

    private static HashSet<string> GetPivotItemNames(Excel.PivotField field)
    {
        Excel.PivotItems? items = null;
        try
        {
            items = (Excel.PivotItems)field.PivotItems();
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var index = 1; index <= items.Count; index++)
            {
                Excel.PivotItem? item = null;
                try
                {
                    item = items.Item(index);
                    names.Add(item.Name);
                }
                finally
                {
                    ComUtilities.Release(ref item);
                }
            }

            return names;
        }
        finally
        {
            ComUtilities.Release(ref items);
        }
    }

    private static void RenameGeneratedGroup(
        Excel.PivotTable pivot,
        string groupedFieldName,
        string generatedGroupItemName,
        string groupName)
    {
        Excel.PivotField? groupedField = null;
        Excel.PivotItem? groupedItem = null;
        try
        {
            groupedField = (Excel.PivotField)pivot.PivotFields(groupedFieldName);
            groupedItem = (Excel.PivotItem)groupedField.PivotItems(generatedGroupItemName);
            groupedItem.Caption = groupName;
        }
        finally
        {
            ComUtilities.Release(ref groupedItem);
            ComUtilities.Release(ref groupedField);
        }
    }
}
