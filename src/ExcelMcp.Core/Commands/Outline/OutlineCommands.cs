using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Core.Commands.Outline;

/// <summary>Excel COM implementation of explicit worksheet row/column outline levels.</summary>
public sealed class OutlineCommands : IOutlineCommands
{
    private const int MaximumPublicLevel = 7;

    /// <inheritdoc />
    public OutlineStateResult SetLevel(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        int level,
        OutlineAxis axis,
        bool? collapsed = null)
    {
        ValidateInput(sheetName, rangeAddress);
        if (level is < 0 or > MaximumPublicLevel)
        {
            throw new ArgumentOutOfRangeException(
                nameof(level),
                level,
                $"level must be between 0 (ungrouped) and {MaximumPublicLevel}");
        }

        return batch.Execute((ctx, ct) =>
        {
            object? sheet = null;
            object? range = null;
            object? units = null;

            try
            {
                sheet = ctx.Book.Worksheets[sheetName];
                dynamic dynamicSheet = sheet;
                range = dynamicSheet.Range[rangeAddress];
                units = ResolveAndValidateUnits(sheet, range, axis);
                var before = ReadStateCore(batch, sheetName, range, units, axis, "set-level", changed: false);

                if (before.MinimumLevel != before.MaximumLevel)
                {
                    throw new InvalidOperationException(
                        $"The requested {axis.ToString().ToLowerInvariant()} range has mixed outline levels " +
                        $"({before.MinimumLevel}-{before.MaximumLevel}); choose a uniform range before setting a level.");
                }

                var currentLevel = before.MinimumLevel;
                var requestedCollapseChange = collapsed.HasValue && before.Collapsed != collapsed;
                var changed = currentLevel != level || requestedCollapseChange;
                dynamic dynamicUnits = units;

                if (currentLevel > level && before.Collapsed == true)
                {
                    dynamicUnits.Hidden = false;
                }

                while (currentLevel < level)
                {
                    dynamicUnits.Group();
                    currentLevel++;
                }

                while (currentLevel > level)
                {
                    dynamicUnits.Ungroup();
                    currentLevel--;
                }

                if (level == 0)
                {
                    dynamicUnits.Hidden = false;
                }
                else if (collapsed.HasValue)
                {
                    // Hidden is the deterministic per-range collapse state used by Excel outlines.
                    // It avoids ActiveCell/Selection and round-trips reliably through COM.
                    dynamicUnits.Hidden = collapsed.Value;
                }

                var result = ReadStateCore(batch, sheetName, range, units, axis, "set-level", changed);
                result.Message = $"Set {result.Axis} outline {result.RangeAddress} to level {result.Level}" +
                    (collapsed.HasValue ? $", collapsed={result.Collapsed}" : string.Empty);
                return result;
            }
            finally
            {
                ComUtilities.Release(ref units);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public OutlineStateResult GetState(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        OutlineAxis axis)
    {
        ValidateInput(sheetName, rangeAddress);

        return batch.Execute((ctx, ct) =>
        {
            object? sheet = null;
            object? range = null;
            object? units = null;

            try
            {
                sheet = ctx.Book.Worksheets[sheetName];
                dynamic dynamicSheet = sheet;
                range = dynamicSheet.Range[rangeAddress];
                units = ResolveAndValidateUnits(sheet, range, axis);
                var result = ReadStateCore(batch, sheetName, range, units, axis, "get-state", changed: false);
                result.Message = result.Level.HasValue
                    ? $"{result.Axis} outline {result.RangeAddress} is level {result.Level}, collapsed={result.Collapsed}"
                    : $"{result.Axis} outline {result.RangeAddress} has mixed levels {result.MinimumLevel}-{result.MaximumLevel}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref units);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static void ValidateInput(string sheetName, string rangeAddress)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            throw new ArgumentException("sheetName is required", nameof(sheetName));
        if (string.IsNullOrWhiteSpace(rangeAddress))
            throw new ArgumentException("rangeAddress is required", nameof(rangeAddress));
    }

    private static object ResolveAndValidateUnits(object sheetObject, object rangeObject, OutlineAxis axis)
    {
        dynamic sheet = sheetObject;
        dynamic range = rangeObject;
        object? rangeRows = null;
        object? rangeColumns = null;
        object? sheetRows = null;
        object? sheetColumns = null;

        try
        {
            rangeRows = range.Rows;
            rangeColumns = range.Columns;
            sheetRows = sheet.Rows;
            sheetColumns = sheet.Columns;
            dynamic dynamicRangeRows = rangeRows;
            dynamic dynamicRangeColumns = rangeColumns;
            dynamic dynamicSheetRows = sheetRows;
            dynamic dynamicSheetColumns = sheetColumns;
            var rangeRowCount = Convert.ToInt32(dynamicRangeRows.Count, CultureInfo.InvariantCulture);
            var rangeColumnCount = Convert.ToInt32(dynamicRangeColumns.Count, CultureInfo.InvariantCulture);
            var sheetRowCount = Convert.ToInt32(dynamicSheetRows.Count, CultureInfo.InvariantCulture);
            var sheetColumnCount = Convert.ToInt32(dynamicSheetColumns.Count, CultureInfo.InvariantCulture);

            if (axis == OutlineAxis.Row && rangeColumnCount != sheetColumnCount)
            {
                throw new ArgumentException(
                    "Row outlines require a complete row range such as 5:10");
            }

            if (axis == OutlineAxis.Column && rangeRowCount != sheetRowCount)
            {
                throw new ArgumentException(
                    "Column outlines require a complete column range such as B:D");
            }

            object result = axis == OutlineAxis.Row ? rangeRows : rangeColumns;
            if (axis == OutlineAxis.Row)
                rangeRows = null;
            else
                rangeColumns = null;
            return result;
        }
        finally
        {
            ComUtilities.Release(ref sheetColumns);
            ComUtilities.Release(ref sheetRows);
            ComUtilities.Release(ref rangeColumns);
            ComUtilities.Release(ref rangeRows);
        }
    }

    private static OutlineStateResult ReadStateCore(
        IExcelBatch batch,
        string sheetName,
        object rangeObject,
        object unitsObject,
        OutlineAxis axis,
        string action,
        bool changed)
    {
        dynamic range = rangeObject;
        dynamic units = unitsObject;
        var count = Convert.ToInt32(units.Count, CultureInfo.InvariantCulture);
        var minimum = int.MaxValue;
        var maximum = int.MinValue;
        bool? collapsed = null;
        var collapseInitialized = false;
        var collapseMixed = false;

        for (var index = 1; index <= count; index++)
        {
            object? unit = null;
            try
            {
                unit = units.Item(index);
                dynamic dynamicUnit = unit;
                var excelLevel = Convert.ToInt32(dynamicUnit.OutlineLevel, CultureInfo.InvariantCulture);
                var publicLevel = Math.Max(0, excelLevel - 1);
                minimum = Math.Min(minimum, publicLevel);
                maximum = Math.Max(maximum, publicLevel);

                var unitCollapsed = Convert.ToBoolean(dynamicUnit.Hidden, CultureInfo.InvariantCulture);
                if (!collapseInitialized)
                {
                    collapsed = unitCollapsed;
                    collapseInitialized = true;
                }
                else if (collapsed != unitCollapsed)
                {
                    collapseMixed = true;
                }
            }
            finally
            {
                ComUtilities.Release(ref unit);
            }
        }

        if (count == 0)
        {
            minimum = 0;
            maximum = 0;
        }

        var normalizedAddress = Convert.ToString(range.Address, CultureInfo.InvariantCulture) ?? string.Empty;
        return new OutlineStateResult
        {
            Success = true,
            FilePath = batch.WorkbookPath,
            Action = action,
            SheetName = sheetName,
            Axis = axis.ToString().ToLowerInvariant(),
            RangeAddress = normalizedAddress,
            Level = minimum == maximum ? minimum : null,
            MinimumLevel = minimum,
            MaximumLevel = maximum,
            Collapsed = collapseMixed ? null : collapsed,
            UnitCount = count,
            Changed = changed,
        };
    }
}
