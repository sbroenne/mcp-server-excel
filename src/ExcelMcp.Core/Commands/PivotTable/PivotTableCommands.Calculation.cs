using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

public partial class PivotTableCommands
{
    // BaseItem accepts these native names. File-format numeric sentinels can crash Excel.
    private const string PreviousBaseItem = "(previous)";
    private const string NextBaseItem = "(next)";

    /// <inheritdoc/>
    public PivotFieldCalculationResult SetFieldCalculation(
        IExcelBatch batch, string pivotTableName, string fieldName,
        PivotFieldCalculation calculation, string? baseFieldName = null,
        PivotCalculationBaseItemKind? baseItemKind = null, string? baseItemName = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(pivotTableName);
        ArgumentException.ThrowIfNullOrWhiteSpace(fieldName);
        if (!Enum.IsDefined(calculation))
            throw new ArgumentOutOfRangeException(nameof(calculation));
        if (baseItemKind.HasValue && !Enum.IsDefined(baseItemKind.Value))
            throw new ArgumentOutOfRangeException(nameof(baseItemKind));
        bool needsField = UsesBaseField(calculation);
        bool needsItem = UsesBaseItem(calculation);
        if (needsField)
            ArgumentException.ThrowIfNullOrWhiteSpace(baseFieldName);
        else if (baseFieldName is not null)
            throw new ArgumentException("This calculation does not use baseFieldName.", nameof(baseFieldName));
        if (needsItem && baseItemKind is null)
            throw new ArgumentException("This calculation requires an explicit baseItemKind.", nameof(baseItemKind));
        if (!needsItem && (baseItemKind is not null || baseItemName is not null))
            throw new ArgumentException("This calculation does not use a base item.", nameof(baseItemKind));
        if (baseItemKind == PivotCalculationBaseItemKind.Named)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(baseItemName);
            if (IsRelativeBaseItemName(baseItemName))
                throw new ArgumentException("Excel reserves '(previous)' and '(next)' for relative base items. Select the corresponding baseItemKind explicitly.", nameof(baseItemName));
        }
        else if (baseItemName is not null)
            throw new ArgumentException("baseItemName is valid only for baseItemKind=Named.", nameof(baseItemName));

        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;
            Excel.PivotCache? cache = null;
            Excel.PivotFields? dataFields = null;
            Excel.PivotField? valueField = null;
            Excel.PivotField? baseField = null;
            try
            {
                pivot = (Excel.PivotTable)FindPivotTable(ctx.Book, pivotTableName);
                cache = pivot.PivotCache();
                bool isOlap = cache.OLAP;
                if (isOlap && needsField)
                    throw new NotSupportedException("Excel does not expose BaseField/BaseItem for OLAP/Data Model PivotTables. Use a source measure for base-dependent calculations.");
                dataFields = (Excel.PivotFields)pivot.DataFields;
                valueField = FindValueField(dataFields, fieldName, ct);
                ValidateCalculationAxis(pivot, calculation);
                if (needsField)
                {
                    baseField = FindCalculationBaseField(pivot, baseFieldName!, ct);
                    if (baseItemKind == PivotCalculationBaseItemKind.Named)
                        ValidateCalculationBaseItem(baseField, baseItemName!, ct);
                }

                ct.ThrowIfCancellationRequested();
                valueField.Calculation = (Excel.XlPivotFieldCalculation)calculation;
                if (baseField is not null)
                    valueField.BaseField = baseField.Name;
                if (needsItem)
                {
                    valueField.BaseItem = baseItemKind switch
                    {
                        PivotCalculationBaseItemKind.Previous => PreviousBaseItem,
                        PivotCalculationBaseItemKind.Next => NextBaseItem,
                        _ => baseItemName!
                    };
                }
                var result = ReadValueFieldCalculation(valueField, isOlap, batch.WorkbookPath);
                if (result.Calculation != calculation ||
                    (needsField && !string.Equals(result.BaseFieldName, baseFieldName, StringComparison.Ordinal)) ||
                    (needsItem && result.BaseItemKind != baseItemKind) ||
                    (baseItemKind == PivotCalculationBaseItemKind.Named &&
                        !string.Equals(result.BaseItemName, baseItemName, StringComparison.Ordinal)))
                {
                    throw new InvalidOperationException(
                        $"Excel did not apply the requested Show Values As settings for '{fieldName}'. " +
                        $"Requested calculation: {calculation}; native calculation: {result.Calculation}. " +
                        (isOlap ? "Use a source measure for this OLAP/Data Model calculation. " : "") +
                        "Read list-fields and displayed data before retrying; failed writes do not promise rollback.");
                }
                return result;
            }
            finally
            {
                ComUtilities.Release(ref baseField);
                ComUtilities.Release(ref valueField);
                ComUtilities.Release(ref dataFields);
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    private static bool UsesBaseItem(PivotFieldCalculation calculation) =>
        calculation is PivotFieldCalculation.DifferenceFrom or PivotFieldCalculation.PercentOf
            or PivotFieldCalculation.PercentDifferenceFrom;

    private static bool UsesBaseField(PivotFieldCalculation calculation) =>
        UsesBaseItem(calculation) || calculation is PivotFieldCalculation.RunningTotal
            or PivotFieldCalculation.PercentRunningTotal or PivotFieldCalculation.PercentOfParent
            or PivotFieldCalculation.RankAscending or PivotFieldCalculation.RankDescending;

    private static bool IsRelativeBaseItemName(string name) =>
        string.Equals(name, PreviousBaseItem, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, NextBaseItem, StringComparison.OrdinalIgnoreCase);

    private static void ValidateCalculationAxis(Excel.PivotTable pivot, PivotFieldCalculation calculation)
    {
        if (calculation is not (PivotFieldCalculation.PercentOfParentRow or PivotFieldCalculation.PercentOfParentColumn))
            return;
        Excel.PivotFields? fields = null;
        try
        {
            fields = (Excel.PivotFields)(calculation == PivotFieldCalculation.PercentOfParentRow
                ? pivot.RowFields : pivot.ColumnFields);
            if (fields.Count == 0)
                throw new ArgumentException("Parent-row/column percentage requires a field on the corresponding axis.", nameof(calculation));
        }
        finally
        {
            ComUtilities.Release(ref fields);
        }
    }

    private static Excel.PivotField FindValueField(Excel.PivotFields fields, string name, CancellationToken ct)
    {
        for (int i = 1; i <= fields.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            Excel.PivotField? field = null;
            try
            {
                field = fields.Item(i);
                if (string.Equals(field.Name, name, StringComparison.Ordinal))
                {
                    var found = field;
                    field = null;
                    return found;
                }
            }
            finally
            {
                ComUtilities.Release(ref field);
            }
        }
        throw new ArgumentException($"Displayed value field '{name}' was not found. Use the exact valueFields.fieldName from list-fields; source names do not identify repeated Values instances.", nameof(name));
    }

    private static Excel.PivotField FindCalculationBaseField(Excel.PivotTable pivot, string name, CancellationToken ct)
    {
        Excel.PivotFields? fields = null;
        try
        {
            fields = (Excel.PivotFields)pivot.PivotFields(Type.Missing);
            for (int i = 1; i <= fields.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.PivotField? field = null;
                try
                {
                    field = fields.Item(i);
                    if (string.Equals(field.Name, name, StringComparison.Ordinal) &&
                        field.Orientation is Excel.XlPivotFieldOrientation.xlRowField or Excel.XlPivotFieldOrientation.xlColumnField)
                    {
                        var found = field;
                        field = null;
                        return found;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref field);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref fields);
        }
        throw new ArgumentException($"Base field '{name}' must be an exact field name placed in Rows or Columns.", nameof(name));
    }

    private static void ValidateCalculationBaseItem(Excel.PivotField field, string name, CancellationToken ct)
    {
        Excel.PivotItems? items = null;
        try
        {
            items = (Excel.PivotItems)field.PivotItems(Type.Missing);
            for (int i = 1; i <= items.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.PivotItem? item = null;
                try
                {
                    item = items.Item(i);
                    if (string.Equals(item.Name, name, StringComparison.Ordinal))
                        return;
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
        }
        throw new ArgumentException($"Base item '{name}' was not found in field '{field.Name}'. Use an exact native item name, not a numeric index.", nameof(name));
    }

    private static PivotFieldCalculationResult ReadValueFieldCalculation(Excel.PivotField field, bool isOlap, string path)
    {
        var calculation = (PivotFieldCalculation)field.Calculation;
        if (!Enum.IsDefined(calculation))
            throw new NotSupportedException($"Native Pivot calculation '{field.Calculation}' is not supported.");
        var result = new PivotFieldCalculationResult
        {
            Success = true,
            FilePath = path,
            FieldName = field.Name,
            SourceName = field.SourceName,
            Position = Convert.ToInt32(field.Position, CultureInfo.InvariantCulture),
            Function = isOlap ? null : PivotTableHelpers.GetAggregationFunctionFromCom((int)field.Function),
            Calculation = calculation,
            IsOlap = isOlap,
            BaseSettingsAvailable = !isOlap
        };
        if (!isOlap && UsesBaseField(calculation))
        {
            object? nativeBase = null;
            try
            {
                nativeBase = field.BaseField;
                result.BaseFieldName = nativeBase is string name ? name :
                    (nativeBase as Excel.PivotField)?.Name ??
                    throw new InvalidOperationException("Excel returned an unsupported BaseField value.");
            }
            finally
            {
                ComUtilities.Release(ref nativeBase);
            }
        }
        if (!isOlap && UsesBaseItem(calculation))
        {
            object? item = null;
            try
            {
                item = field.BaseItem;
                var name = item is string text ? text :
                    (item as Excel.PivotItem)?.Name ??
                    throw new InvalidOperationException("Excel returned an unsupported BaseItem value.");
                if (string.Equals(name, PreviousBaseItem, StringComparison.OrdinalIgnoreCase))
                {
                    result.BaseItemKind = PivotCalculationBaseItemKind.Previous;
                }
                else if (string.Equals(name, NextBaseItem, StringComparison.OrdinalIgnoreCase))
                {
                    result.BaseItemKind = PivotCalculationBaseItemKind.Next;
                }
                else
                {
                    result.BaseItemKind = PivotCalculationBaseItemKind.Named;
                    result.BaseItemName = name;
                }
            }
            finally
            {
                ComUtilities.Release(ref item);
            }
        }
        return result;
    }

    private static List<PivotFieldCalculationResult> ReadAllValueFieldCalculations(
        Excel.PivotTable pivot, string path, CancellationToken ct)
    {
        Excel.PivotCache? cache = null;
        Excel.PivotFields? fields = null;
        try
        {
            cache = pivot.PivotCache();
            fields = (Excel.PivotFields)pivot.DataFields;
            bool isOlap = cache.OLAP;
            var result = new List<PivotFieldCalculationResult>();
            for (int i = 1; i <= fields.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.PivotField? field = null;
                try
                {
                    field = fields.Item(i);
                    result.Add(ReadValueFieldCalculation(field, isOlap, path));
                }
                finally
                {
                    ComUtilities.Release(ref field);
                }
            }
            return result;
        }
        finally
        {
            ComUtilities.Release(ref fields);
            ComUtilities.Release(ref cache);
        }
    }
}
