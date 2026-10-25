using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

public partial class PivotTableCommands
{
    /// <inheritdoc/>
    public PivotFieldOptionsResult GetFieldOptions(IExcelBatch batch, string pivotTableName, string fieldName) =>
        FieldOptions(batch, pivotTableName, fieldName, null);

    /// <inheritdoc/>
    public PivotFieldOptionsResult SetFieldOptions(IExcelBatch batch, string pivotTableName, string fieldName, bool showAllItems) =>
        FieldOptions(batch, pivotTableName, fieldName, showAllItems);

    private static PivotFieldOptionsResult FieldOptions(IExcelBatch batch, string name, string fieldName, bool? showAllItems) =>
        WithRegularPlacedField(batch, name, fieldName, (_, field, ct) =>
        {
            RequireAxisField(field);
            ct.ThrowIfCancellationRequested();
            if (showAllItems.HasValue)
                field.ShowAllItems = showAllItems.Value;
            bool actual = field.ShowAllItems;
            if (showAllItems.HasValue && actual != showAllItems.Value)
                throw new InvalidOperationException("Excel did not apply showAllItems. Inspect get-field-options before retrying.");
            return new PivotFieldOptionsResult
            {
                Success = true,
                FilePath = batch.WorkbookPath,
                PivotTableName = name,
                FieldName = field.Name,
                ShowAllItems = actual
            };
        }, mutation: showAllItems.HasValue);

    /// <inheritdoc/>
    public PivotSubtotalResult GetSubtotalFunctions(IExcelBatch batch, string pivotTableName, string fieldName) =>
        WithRegularPlacedField(batch, pivotTableName, fieldName, (_, field, ct) =>
        {
            RequireAxisField(field);
            return ReadSubtotalFunctions(field, pivotTableName, batch.WorkbookPath, ct);
        });

    /// <inheritdoc/>
    public PivotSubtotalResult SetSubtotalFunctions(IExcelBatch batch, string pivotTableName, string fieldName,
        List<PivotSubtotalFunction> functions)
    {
        ArgumentNullException.ThrowIfNull(functions);
        if (functions.Any(function => !Enum.IsDefined(function)) || functions.Distinct().Count() != functions.Count ||
            (functions.Contains(PivotSubtotalFunction.Automatic) && functions.Count != 1))
            throw new ArgumentException("Use distinct valid subtotal functions; Automatic cannot be combined with other functions.", nameof(functions));
        return WithRegularPlacedField(batch, pivotTableName, fieldName, (_, field, ct) =>
        {
            RequireAxisField(field);
            ct.ThrowIfCancellationRequested();
            object[] flags = Enumerable.Range(1, 12)
                .Select(slot => (object)functions.Contains((PivotSubtotalFunction)slot)).ToArray();
            field.Subtotals[Type.Missing] = flags;
            var result = ReadSubtotalFunctions(field, pivotTableName, batch.WorkbookPath, ct);
            if (!result.Functions.SequenceEqual(functions.Order()))
                throw new InvalidOperationException("Excel did not apply all requested subtotals. Inspect get-subtotal-functions before retrying.");
            return result;
        }, mutation: true);
    }

    private static PivotSubtotalResult ReadSubtotalFunctions(Excel.PivotField field, string name, string path, CancellationToken ct)
    {
        var result = new PivotSubtotalResult
        {
            Success = true,
            FilePath = path,
            PivotTableName = name,
            FieldName = field.Name
        };
        for (int slot = 1; slot <= 12; slot++)
        {
            ct.ThrowIfCancellationRequested();
            if (Convert.ToBoolean((object)field.Subtotals[slot], System.Globalization.CultureInfo.InvariantCulture))
                result.Functions.Add((PivotSubtotalFunction)slot);
        }
        return result;
    }

    /// <inheritdoc/>
    public PivotCalculatedItemsResult CreateCalculatedItem(IExcelBatch batch, string pivotTableName,
        string fieldName, string itemName, string formula)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(itemName);
        ArgumentException.ThrowIfNullOrWhiteSpace(formula);
        return CalculatedItems(batch, pivotTableName, fieldName, itemName, formula);
    }

    /// <inheritdoc/>
    public PivotCalculatedItemsResult ListCalculatedItems(IExcelBatch batch, string pivotTableName, string fieldName) =>
        CalculatedItems(batch, pivotTableName, fieldName, null, null);

    /// <inheritdoc/>
    public PivotCalculatedItemsResult DeleteCalculatedItem(IExcelBatch batch, string pivotTableName,
        string fieldName, string itemName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(itemName);
        return CalculatedItems(batch, pivotTableName, fieldName, itemName, null);
    }

    private static PivotCalculatedItemsResult CalculatedItems(IExcelBatch batch, string name, string fieldName,
        string? itemName, string? formula) =>
        WithRegularPlacedField(batch, name, fieldName, (pivot, field, ct) =>
        {
            RequireAxisField(field);
            Excel.CalculatedItems? calculated = null;
            Excel.PivotItems? allItems = null;
            Excel.PivotItem? item = null;
            Excel.PivotFields? fields = null;
            try
            {
                if (itemName is not null)
                {
                    fields = (Excel.PivotFields)pivot.PivotFields();
                    for (int index = 1; index <= fields.Count; index++)
                    {
                        ct.ThrowIfCancellationRequested();
                        Excel.PivotField? candidate = null;
                        try
                        {
                            candidate = fields.Item(index);
                            if (candidate.GroupLevel > 1)
                                throw new NotSupportedException("Calculated-item writes are not supported in PivotTables with grouped fields. Ungroup first.");
                        }
                        finally
                        {
                            ComUtilities.Release(ref candidate);
                        }
                    }
                }
                calculated = field.CalculatedItems();
                if (formula is not null)
                {
                    allItems = (Excel.PivotItems)field.PivotItems();
                    for (int index = 1; index <= allItems.Count; index++)
                    {
                        ct.ThrowIfCancellationRequested();
                        item = allItems.Item(index);
                        if (string.Equals(item.Name, itemName, StringComparison.OrdinalIgnoreCase))
                            throw new ArgumentException($"Item '{itemName}' already exists; choose a new name.");
                        ComUtilities.Release(ref item);
                    }
                    ct.ThrowIfCancellationRequested();
                    item = calculated.Add(itemName!, formula, true);
                    ComUtilities.Release(ref item);
                }
                else if (itemName is not null)
                {
                    for (int index = 1; index <= calculated.Count; index++)
                    {
                        ct.ThrowIfCancellationRequested();
                        item = calculated.Item(index);
                        if (string.Equals(item.Name, itemName, StringComparison.Ordinal))
                            break;
                        ComUtilities.Release(ref item);
                    }
                    if (item is null)
                        throw new ArgumentException($"Calculated item '{itemName}' was not found. Use list-calculated-items; source items cannot be deleted by this action.");
                    ct.ThrowIfCancellationRequested();
                    item.Delete();
                    ComUtilities.Release(ref item);
                }
                var result = new PivotCalculatedItemsResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    PivotTableName = name,
                    FieldName = field.Name
                };
                for (int index = 1; index <= calculated.Count; index++)
                {
                    ct.ThrowIfCancellationRequested();
                    try
                    {
                        item = calculated.Item(index);
                        result.Items.Add(new PivotCalculatedItemInfo { Name = item.Name, Formula = item.Formula });
                    }
                    finally
                    {
                        ComUtilities.Release(ref item);
                    }
                }
                if (itemName is not null && result.Items.Any(value => value.Name == itemName) != (formula is not null))
                    throw new InvalidOperationException("Excel did not apply the calculated-item change. Inspect list-calculated-items and get-data before retrying.");
                return result;
            }
            finally
            {
                ComUtilities.Release(ref item);
                ComUtilities.Release(ref allItems);
                ComUtilities.Release(ref calculated);
                ComUtilities.Release(ref fields);
            }
        }, mutation: itemName is not null);

    private static void RequireAxisField(Excel.PivotField field)
    {
        if (field.Orientation is not (Excel.XlPivotFieldOrientation.xlRowField or Excel.XlPivotFieldOrientation.xlColumnField))
            throw new ArgumentException("Select the exact caption of a placed row or column field.");
    }
}
