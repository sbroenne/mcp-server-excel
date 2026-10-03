using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

public partial class PivotTableCommands
{
    /// <inheritdoc/>
    public PivotFiltersResult GetFieldFilters(IExcelBatch batch, string pivotTableName, string fieldName) =>
        WithRegularPlacedField(batch, pivotTableName, fieldName,
            (_, field, ct) => ReadPivotFilters(field, pivotTableName, batch.WorkbookPath, ct));

    /// <inheritdoc/>
    public PivotFiltersResult AddFieldFilter(IExcelBatch batch, string pivotTableName, string fieldName,
        PivotFilterOptions filterOptions)
    {
        ArgumentNullException.ThrowIfNull(filterOptions);
        var (type, value1, value2) = ValidatePivotFilter(filterOptions);
        return WithRegularPlacedField(batch, pivotTableName, fieldName, (pivot, field, ct) =>
        {
            Excel.PivotFields? dataFields = null;
            Excel.PivotField? dataField = null;
            Excel.PivotFilters? filters = null;
            Excel.PivotFilter? added = null;
            try
            {
                if (filterOptions.DataFieldName is not null)
                {
                    dataFields = (Excel.PivotFields)pivot.DataFields;
                    dataField = FindValueField(dataFields, filterOptions.DataFieldName, ct);
                }
                if (filterOptions.Date1.HasValue && field.DataType != Excel.XlPivotFieldDataType.xlDate)
                    throw new ArgumentException("Date filters require a native date field; grouping a text label does not make it a date.");
                filters = field.PivotFilters;
                if (filters.Count > 0 && !pivot.AllowMultipleFilters)
                    throw new InvalidOperationException("This field already has a calculated filter. Clear it explicitly or set layoutOptions.allowMultipleFilters=true before adding another.");
                ct.ThrowIfCancellationRequested();
                added = filters.Add2(type, dataField ?? (object)Type.Missing, value1, value2,
                    WholeDayFilter: filterOptions.Date1.HasValue ? true : Type.Missing);
                return ReadPivotFilters(field, pivotTableName, batch.WorkbookPath, ct);
            }
            finally
            {
                ComUtilities.Release(ref added);
                ComUtilities.Release(ref filters);
                ComUtilities.Release(ref dataField);
                ComUtilities.Release(ref dataFields);
            }
        }, mutation: true);
    }

    /// <inheritdoc/>
    public PivotFiltersResult ClearFieldFilters(IExcelBatch batch, string pivotTableName, string fieldName) =>
        WithRegularPlacedField(batch, pivotTableName, fieldName, (_, field, ct) =>
        {
            Excel.PivotFilters? filters = null;
            try
            {
                filters = field.PivotFilters;
                for (int index = filters.Count; index >= 1; index--)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.PivotFilter? filter = null;
                    try
                    {
                        filter = filters[index];
                        filter.Delete();
                    }
                    finally
                    {
                        ComUtilities.Release(ref filter);
                    }
                }
                return ReadPivotFilters(field, pivotTableName, batch.WorkbookPath, ct);
            }
            finally
            {
                ComUtilities.Release(ref filters);
            }
        }, mutation: true);

    /// <inheritdoc/>
    public PivotItemExpansionResult GetItemExpansion(IExcelBatch batch, string pivotTableName, string fieldName,
        string itemName) => ItemExpansion(batch, pivotTableName, fieldName, itemName, null);

    /// <inheritdoc/>
    public PivotItemExpansionResult SetItemExpansion(IExcelBatch batch, string pivotTableName, string fieldName,
        string itemName, bool expanded) => ItemExpansion(batch, pivotTableName, fieldName, itemName, expanded);

    private static PivotItemExpansionResult ItemExpansion(IExcelBatch batch, string pivotTableName,
        string fieldName, string itemName, bool? expanded)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(itemName);
        return WithRegularPlacedField(batch, pivotTableName, fieldName, (pivot, field, ct) =>
        {
            Excel.PivotFields? axis = null;
            Excel.PivotItems? items = null;
            Excel.PivotItem? item = null;
            try
            {
                if (field.Orientation is not (Excel.XlPivotFieldOrientation.xlRowField or Excel.XlPivotFieldOrientation.xlColumnField))
                    throw new ArgumentException("Expansion requires a placed row or column field.");
                axis = (Excel.PivotFields)(field.Orientation == Excel.XlPivotFieldOrientation.xlRowField
                    ? pivot.RowFields : pivot.ColumnFields);
                if (field.Position >= axis.Count)
                    throw new ArgumentException("An innermost field has no child level to expand or collapse.");
                items = (Excel.PivotItems)field.PivotItems();
                for (int index = 1; index <= items.Count; index++)
                {
                    ct.ThrowIfCancellationRequested();
                    item = items.Item(index);
                    if (string.Equals(item.Name, itemName, StringComparison.Ordinal))
                        break;
                    ComUtilities.Release(ref item);
                }
                if (item is null)
                    throw new ArgumentException($"Item '{itemName}' was not found; use its exact native caption.");
                if (!item.Visible)
                    throw new ArgumentException("A hidden item cannot be expanded or collapsed.");
                if (expanded.HasValue)
                    item.ShowDetail = expanded.Value;
                return new PivotItemExpansionResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    PivotTableName = pivotTableName,
                    FieldName = field.Name,
                    ItemName = item.Name,
                    Expanded = item.ShowDetail
                };
            }
            finally
            {
                ComUtilities.Release(ref item);
                ComUtilities.Release(ref items);
                ComUtilities.Release(ref axis);
            }
        }, mutation: expanded.HasValue);
    }

    private static TResult WithRegularPlacedField<TResult>(IExcelBatch batch, string pivotTableName,
        string fieldName, Func<Excel.PivotTable, Excel.PivotField, CancellationToken, TResult> action, bool mutation = false)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(pivotTableName);
        ArgumentException.ThrowIfNullOrWhiteSpace(fieldName);
        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;
            Excel.PivotCache? cache = null;
            Excel.PivotFields? fields = null;
            Excel.PivotField? field = null;
            Excel.Worksheet? sheet = null;
            try
            {
                pivot = (Excel.PivotTable)FindPivotTable(ctx.Book, pivotTableName);
                cache = pivot.PivotCache();
                if (cache.OLAP)
                    throw new NotSupportedException("These native field-filter/expansion operations support regular PivotTables only. OLAP/Data Model filtering and hierarchy expansion are provider-dependent.");
                sheet = (Excel.Worksheet)pivot.Parent;
                if (mutation && sheet.ProtectContents)
                    throw new InvalidOperationException("Unprotect the PivotTable worksheet before using this field operation.");
                fields = (Excel.PivotFields)pivot.PivotFields();
                for (int index = 1; index <= fields.Count; index++)
                {
                    ct.ThrowIfCancellationRequested();
                    field = fields.Item(index);
                    if (string.Equals(field.Name, fieldName, StringComparison.Ordinal))
                        break;
                    ComUtilities.Release(ref field);
                }
                if (field is null || field.Orientation is Excel.XlPivotFieldOrientation.xlHidden or Excel.XlPivotFieldOrientation.xlDataField)
                    throw new ArgumentException("Select the exact caption of a placed row, column, or page field.");
                return action(pivot, field, ct);
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref fields);
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    private static (Excel.XlPivotFilterType Type, object Value1, object Value2) ValidatePivotFilter(PivotFilterOptions options)
    {
        if (!Enum.IsDefined(options.Type))
            throw new ArgumentOutOfRangeException(nameof(options));
        string name = options.Type.ToString();
        var type = Enum.Parse<Excel.XlPivotFilterType>("xl" + name);
        bool label = name.StartsWith("Caption", StringComparison.Ordinal);
        bool date = IsDateComparison(name);
        bool top = name.StartsWith("Top", StringComparison.Ordinal) || name.StartsWith("Bottom", StringComparison.Ordinal);
        bool between = name.Contains("Between", StringComparison.Ordinal);
        if (label)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(options.Text1);
            if (between)
                ArgumentException.ThrowIfNullOrWhiteSpace(options.Text2);
            if (options.Number1.HasValue || options.Number2.HasValue || options.Date1.HasValue ||
                options.Date2.HasValue || options.DataFieldName is not null || (!between && options.Text2 is not null))
                throw new ArgumentException("Label filters accept only text1 and, for intervals, text2.");
            return (type, options.Text1, options.Text2 ?? (object)Type.Missing);
        }
        if (date)
        {
            if (!options.Date1.HasValue || (between && !options.Date2.HasValue))
                throw new ArgumentException("Date filters require date1 and, for intervals, date2.");
            if (options.Text1 is not null || options.Text2 is not null || options.Number1.HasValue ||
                options.Number2.HasValue || options.DataFieldName is not null || (!between && options.Date2.HasValue))
                throw new ArgumentException("Date filters accept only date1 and, for intervals, date2.");
            if (options.Date1.Value < DateTime.FromOADate(0) || options.Date1.Value.Year >= 10000 ||
                (options.Date2.HasValue && options.Date2.Value < options.Date1.Value))
                throw new ArgumentException("Date criteria must be Excel dates in ascending order.");
            return (type, options.Date1.Value.ToOADate(), options.Date2.HasValue ? options.Date2.Value.ToOADate() : Type.Missing);
        }
        ArgumentException.ThrowIfNullOrWhiteSpace(options.DataFieldName);
        if (!options.Number1.HasValue || !double.IsFinite(options.Number1.Value) ||
            (between && (!options.Number2.HasValue || !double.IsFinite(options.Number2.Value) || options.Number2 < options.Number1)))
            throw new ArgumentException("Value filters require finite number1 and ascending number2 for intervals.");
        if (options.Text1 is not null || options.Text2 is not null || options.Date1.HasValue ||
            options.Date2.HasValue || (!between && options.Number2.HasValue))
            throw new ArgumentException("Value/top filters accept only numeric criteria and dataFieldName.");
        if (top && (options.Number1 <= 0 ||
            (name.EndsWith("Count", StringComparison.Ordinal) && options.Number1 != Math.Truncate(options.Number1.Value)) ||
            (name.EndsWith("Percent", StringComparison.Ordinal) && options.Number1 > 100)))
            throw new ArgumentException("Top/bottom amount must be positive; counts must be whole numbers and percentages at most 100.");
        return (type, options.Number1.Value, options.Number2 ?? (object)Type.Missing);
    }

    private static PivotFiltersResult ReadPivotFilters(Excel.PivotField field, string name, string filePath, CancellationToken ct)
    {
        var result = new PivotFiltersResult { Success = true, FilePath = filePath, PivotTableName = name, FieldName = field.Name };
        Excel.PivotFilters? filters = null;
        try
        {
            filters = field.PivotFilters;
            for (int index = 1; index <= filters.Count; index++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.PivotFilter? filter = null;
                Excel.PivotField? dataField = null;
                try
                {
                    filter = filters[index];
                    string type = filter.FilterType.ToString();
                    type = type.StartsWith("xl", StringComparison.Ordinal) ? type[2..] : type;
                    bool criterion = type.StartsWith("Caption", StringComparison.Ordinal) ||
                        type.StartsWith("Value", StringComparison.Ordinal) ||
                        IsDateComparison(type) ||
                        type.StartsWith("Top", StringComparison.Ordinal) || type.StartsWith("Bottom", StringComparison.Ordinal);
                    if (filter.IsMemberPropertyFilter)
                        throw new NotSupportedException("Member-property filters cannot be read as regular caption/value criteria.");
                    if (type.StartsWith("Value", StringComparison.Ordinal) ||
                        type.StartsWith("Top", StringComparison.Ordinal) || type.StartsWith("Bottom", StringComparison.Ordinal))
                        dataField = filter.DataField;
                    result.Filters.Add(new PivotFilterInfo
                    {
                        Index = index,
                        Type = type,
                        Value1 = criterion ? NormalizePivotCriterion(filter.Value1) : null,
                        Value2 = type.Contains("Between", StringComparison.Ordinal) ? NormalizePivotCriterion(filter.Value2) : null,
                        DataFieldName = dataField?.Name
                    });
                }
                finally
                {
                    ComUtilities.Release(ref dataField);
                    ComUtilities.Release(ref filter);
                }
            }
            return result;
        }
        finally
        {
            ComUtilities.Release(ref filters);
        }
    }

    private static object? NormalizePivotCriterion(object? value) =>
        value is DateTime date ? date.ToString("O", CultureInfo.InvariantCulture) : value;

    private static bool IsDateComparison(string type) => type is
        "SpecificDate" or "NotSpecificDate" or "Before" or "BeforeOrEqualTo" or
        "After" or "AfterOrEqualTo" or "DateBetween" or "DateNotBetween";
}
