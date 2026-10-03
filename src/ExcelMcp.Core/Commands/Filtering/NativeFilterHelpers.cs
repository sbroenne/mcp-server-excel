using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Filtering;

internal static class NativeFilterHelpers
{
    internal static void Validate(FilterOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        if (!Enum.IsDefined(options.FilterOperator))
            throw new ArgumentException("Unknown filterOperator.", nameof(options));
        var kind = options.FilterOperator;
        bool comparison = kind is FilterOperator.Comparison or FilterOperator.And or FilterOperator.Or;
        bool pair = kind is FilterOperator.And or FilterOperator.Or;
        bool top = kind is FilterOperator.TopItems or FilterOperator.BottomItems or FilterOperator.TopPercent or FilterOperator.BottomPercent;
        bool color = kind is FilterOperator.CellColor or FilterOperator.FontColor;
        if (comparison != (options.Criteria1 is not null) || pair != (options.Criteria2 is not null) ||
            top != options.Count.HasValue || color != (options.Color is not null) ||
            (kind == FilterOperator.Icon) != (options.IconSet is not null && options.IconIndex.HasValue) ||
            (kind != FilterOperator.Icon && (options.IconSet is not null || options.IconIndex.HasValue)) ||
            (kind == FilterOperator.Dynamic) != (options.DynamicCriteria is not null))
            throw new ArgumentException("Filter settings do not match filterOperator. Supply only its applicable criteria.", nameof(options));
        if (kind == FilterOperator.Values)
        {
            bool values = options.Values is { Count: > 0 };
            bool dates = options.DateGroups is { Count: > 0 };
            if (values == dates || (options.Values is not null && options.DateGroups is not null))
                throw new ArgumentException("Values requires either nonempty values or nonempty dateGroups, not both.", nameof(options));
        }
        else if (options.Values is not null || options.DateGroups is not null)
            throw new ArgumentException("values and dateGroups require filterOperator Values.", nameof(options));
        if (options.Values?.Any(value => value is null) == true)
            throw new ArgumentException("Filter values cannot contain null.", nameof(options));
        if (options.DateGroups?.Any(group => group is null || !Enum.IsDefined(group.Level) ||
            group.Date.Year < 1900) == true)
            throw new ArgumentException("Date groups require native levels and dates from 1900 onward.", nameof(options));
        if (top && (options.Count < 1 ||
            (kind is FilterOperator.TopPercent or FilterOperator.BottomPercent && options.Count > 100)))
            throw new ArgumentException("Top/bottom counts must be positive; percentages cannot exceed 100.", nameof(options));
        if (color)
            _ = FormattingHelpers.ParseColor(options.Color!);
        if (kind == FilterOperator.Icon &&
            (!Enum.TryParse<Excel.XlIconSet>(options.IconSet, true, out var iconSet) || !Enum.IsDefined(iconSet) || options.IconIndex < 1))
            throw new ArgumentException("Icon requires a native XlIconSet name and a positive iconIndex.", nameof(options));
        if (kind == FilterOperator.Dynamic &&
            (!Enum.TryParse<Excel.XlDynamicFilterCriteria>(options.DynamicCriteria, true, out var dynamicCriterion) || !Enum.IsDefined(dynamicCriterion)))
            throw new ArgumentException("Dynamic requires a native XlDynamicFilterCriteria name.", nameof(options));
    }

    internal static void Apply(Excel.Workbook workbook, Excel.Range range, int columnIndex,
        FilterOptions options, CancellationToken token)
    {
        Excel.IconSets? sets = null;
        Excel.IconSet? set = null;
        Excel.Icon? icon = null;
        object criteria1 = Type.Missing;
        object criteria2 = Type.Missing;
        try
        {
            token.ThrowIfCancellationRequested();
            var kind = options.FilterOperator;
            if (options.Criteria1 is not null)
                criteria1 = options.Criteria1;
            if (options.Criteria2 is not null)
                criteria2 = options.Criteria2;
            if (options.Values is not null)
                criteria1 = options.Values.ToArray();
            if (options.DateGroups is not null)
                criteria2 = options.DateGroups.SelectMany(group => new object[]
                {
                    (int)group.Level, group.Date.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture)
                }).ToArray();
            if (options.Count.HasValue)
                criteria1 = options.Count.Value.ToString(CultureInfo.InvariantCulture);
            if (options.Color is not null)
                criteria1 = FormattingHelpers.ParseColor(options.Color);
            if (kind == FilterOperator.Dynamic)
                criteria1 = (int)Enum.Parse<Excel.XlDynamicFilterCriteria>(options.DynamicCriteria!, true);
            if (kind == FilterOperator.Icon)
            {
                sets = workbook.IconSets;
                set = sets[Enum.Parse<Excel.XlIconSet>(options.IconSet!, true)];
                if (options.IconIndex > set.Count)
                    throw new ArgumentException("iconIndex is outside the selected native icon set.");
                icon = set[options.IconIndex!.Value];
                criteria1 = icon;
            }
            if (kind == FilterOperator.Comparison)
                range.AutoFilter(Field: columnIndex, Criteria1: criteria1, VisibleDropDown: true);
            else
                range.AutoFilter(Field: columnIndex, Criteria1: criteria1,
                    Operator: (Excel.XlAutoFilterOperator)(int)kind, Criteria2: criteria2, VisibleDropDown: true);
        }
        finally
        {
            ComUtilities.Release(ref icon);
            ComUtilities.Release(ref set);
            ComUtilities.Release(ref sets);
        }
    }

    internal static List<ColumnFilter> Read(Excel.AutoFilter filter, IReadOnlyList<string> columnNames, CancellationToken token)
    {
        Excel.Filters? filters = null;
        try
        {
            filters = filter.Filters;
            if (filters.Count != columnNames.Count)
                throw new InvalidOperationException("Native filter column count does not match the requested scope.");
            var result = new List<ColumnFilter>(filters.Count);
            for (int index = 1; index <= filters.Count; index++)
            {
                token.ThrowIfCancellationRequested();
                Excel.Filter? column = null;
                try
                {
                    column = filters[index];
                    bool active = column.On;
                    result.Add(new ColumnFilter
                    {
                        ColumnName = columnNames[index - 1],
                        ColumnIndex = index,
                        IsFiltered = active,
                        FilterOperator = active ? (FilterOperator)Convert.ToInt32(column.Operator, CultureInfo.InvariantCulture) : null,
                        Criteria1 = active ? ReadCriterion(() => column.Criteria1) : null,
                        Criteria2 = active ? ReadCriterion(() => column.Criteria2) : null
                    });
                }
                finally
                {
                    ComUtilities.Release(ref column);
                }
            }
            return result;
        }
        finally
        {
            ComUtilities.Release(ref filters);
        }
    }

    private static FilterCriterion ReadCriterion(Func<object> getter)
    {
        object? native = null;
        try
        {
            native = getter();
            return new FilterCriterion { Available = true, EmptyVariant = native is DBNull, Value = ReadValue(native) };
        }
        catch (COMException exception) when (exception.HResult == unchecked((int)0x800A03EC))
        {
            return new FilterCriterion
            {
                Available = false,
                HResult = exception.HResult,
                ReadError = "Excel's criterion getter failed: " + exception.Message
            };
        }
        finally
        {
            ComUtilities.Release(ref native);
        }
    }

    private static object? ReadValue(object? value)
    {
        if (value is Excel.Interior interior)
            return Convert.ToInt32(interior.Color, CultureInfo.InvariantCulture);
        if (value is Excel.Icon icon)
        {
            Excel.IconSet? parent = null;
            try
            {
                parent = (Excel.IconSet)icon.Parent;
                return new FilterIcon { IconSet = parent.ID.ToString(), IconIndex = icon.Index };
            }
            finally
            {
                ComUtilities.Release(ref parent);
            }
        }
        if (value is Array array)
        {
            if (array.Rank != 1)
                throw new InvalidOperationException("Excel returned an unexpected multidimensional filter criterion.");
            var items = new List<object?>(array.Length);
            foreach (object? item in array)
                items.Add(ReadValue(item));
            return items;
        }
        if (value is null or string or bool or double or int or long or short or float or decimal or DateTime)
            return value;
        if (value is DBNull)
            return null;
        throw new InvalidOperationException("Excel returned an unsupported native filter criterion type: " + value.GetType().FullName);
    }
}
