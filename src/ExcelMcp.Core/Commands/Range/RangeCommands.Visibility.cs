using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

public partial class RangeCommands
{
    /// <inheritdoc />
    public RangeVisibilityResult GetVisibility(
        IExcelBatch batch, string sheetName, string rangeAddress, VisibilityAxis axis)
    {
        ValidateVisibilityInputs(rangeAddress, axis);
        return batch.Execute((context, ct) =>
        {
            Excel.Range? scope = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? dimensions = null;
            Excel.AutoFilter? filter = null;
            Excel.Range? filterRange = null;
            Excel.Range? filterRows = null;
            try
            {
                ct.ThrowIfCancellationRequested();
                scope = RangeHelpers.ResolveRange(context.Book, sheetName, rangeAddress, out string? error);
                if (scope is null)
                    throw new InvalidOperationException(error ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                sheet = scope.Worksheet;
                var result = new RangeVisibilityResult
                {
                    FilePath = batch.WorkbookPath,
                    Action = "get-visibility",
                    SheetName = sheet.Name,
                    RangeAddress = scope.Address,
                    Axis = axis,
                    SheetFilterMode = sheet.FilterMode
                };
                int firstFilterDataRow = 0;
                int lastFilterDataRow = -1;
                if (axis == VisibilityAxis.Rows && sheet.AutoFilterMode)
                {
                    filter = sheet.AutoFilter;
                    if (filter is not null)
                    {
                        filterRange = filter.Range;
                        filterRows = filterRange.Rows;
                        firstFilterDataRow = filterRange.Row + 1;
                        lastFilterDataRow = filterRange.Row + filterRows.Count - 1;
                    }
                }
                dimensions = axis == VisibilityAxis.Rows ? sheet.Rows : sheet.Columns;
                foreach (int index in GetVisibilityIndices(scope, axis, ct))
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.Range? dimension = null;
                    try
                    {
                        dimension = dimensions[index];
                        object hidden = dimension.Hidden;
                        if (hidden is not bool isHidden)
                            throw new InvalidOperationException($"Excel returned an indeterminate hidden state for {axis} {index}.");
                        object? nativeSize = axis == VisibilityAxis.Rows ? dimension.RowHeight : dimension.ColumnWidth;
                        result.Items.Add(new DimensionVisibility(index, isHidden,
                            isHidden ? "undetermined" : "not-hidden",
                            nativeSize is null ? null : Convert.ToDouble(nativeSize, CultureInfo.InvariantCulture),
                            axis == VisibilityAxis.Rows ? "points" : "character-width",
                            Convert.ToInt32((object)dimension.OutlineLevel, CultureInfo.InvariantCulture),
                            axis == VisibilityAxis.Rows && index >= firstFilterDataRow && index <= lastFilterDataRow));
                    }
                    finally
                    {
                        ComUtilities.Release(ref dimension);
                    }
                }
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref dimensions);
                ComUtilities.Release(ref filterRows);
                ComUtilities.Release(ref filterRange);
                ComUtilities.Release(ref filter);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref scope);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetVisibility(IExcelBatch batch, string sheetName, string rangeAddress,
        VisibilityAxis axis, bool hidden)
    {
        ValidateVisibilityInputs(rangeAddress, axis);
        return batch.Execute((context, ct) =>
        {
            Excel.Range? scope = null;
            Excel.Range? dimensions = null;
            try
            {
                ct.ThrowIfCancellationRequested();
                scope = RangeHelpers.ResolveRange(context.Book, sheetName, rangeAddress, out string? error);
                if (scope is null)
                    throw new InvalidOperationException(error ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                dimensions = axis == VisibilityAxis.Rows ? scope.EntireRow : scope.EntireColumn;
                ct.ThrowIfCancellationRequested();
                dimensions.Hidden = hidden;
                return new OperationResult
                {
                    FilePath = batch.WorkbookPath,
                    Action = "set-visibility",
                    Success = true,
                    Message = "Whole dimensions updated; filter criteria and outline groups are unchanged."
                };
            }
            finally
            {
                ComUtilities.Release(ref dimensions);
                ComUtilities.Release(ref scope);
            }
        });
    }

    private static void ValidateVisibilityInputs(string rangeAddress, VisibilityAxis axis)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(rangeAddress);
        if (!Enum.IsDefined(axis))
            throw new ArgumentOutOfRangeException(nameof(axis));
    }

    private static SortedSet<int> GetVisibilityIndices(Excel.Range scope, VisibilityAxis axis, CancellationToken ct)
    {
        SortedSet<int> indices = [];
        Excel.Areas? areas = null;
        try
        {
            areas = scope.Areas;
            for (int index = 1; index <= areas.Count; index++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.Range? area = null;
                Excel.Range? dimensions = null;
                try
                {
                    area = areas[index];
                    dimensions = axis == VisibilityAxis.Rows ? area.Rows : area.Columns;
                    int first = axis == VisibilityAxis.Rows ? area.Row : area.Column;
                    int count = dimensions.Count;
                    for (int offset = 0; offset < count; offset++)
                    {
                        ct.ThrowIfCancellationRequested();
                        indices.Add(first + offset);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref dimensions);
                    ComUtilities.Release(ref area);
                }
            }
            return indices;
        }
        finally
        {
            ComUtilities.Release(ref areas);
        }
    }
}
