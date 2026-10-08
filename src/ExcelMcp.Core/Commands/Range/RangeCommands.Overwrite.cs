using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

public partial class RangeCommands
{
    private const int InspectionBlockCells = 16_384;
    private const int ConflictExampleLimit = RangeCommandValidation.ConflictExampleLimit;

    private static void ValidateOverwritePolicy(OverwritePolicy overwritePolicy)
    {
        RangeCommandValidation.ValidateOverwritePolicy(overwritePolicy);
    }

    private static (int Rows, int Columns) GetContentDimensions(Excel.Range range)
    {
        Excel.Areas? areas = null;
        Excel.Range? rows = null;
        Excel.Range? columns = null;
        try
        {
            areas = range.Areas;
            if (areas.Count != 1)
            {
                throw new OperationFailureException(
                    OperationFailureCategory.InvalidInput,
                    "Content writes require a single rectangular range; use separate requests for disjoint ranges.");
            }

            rows = range.Rows;
            columns = range.Columns;
            return (rows.Count, columns.Count);
        }
        finally
        {
            ComUtilities.Release(ref columns);
            ComUtilities.Release(ref rows);
            ComUtilities.Release(ref areas);
        }
    }

    private static void ValidateContentWriteDimensions<T>(
        Excel.Range range, List<List<T>> payload, string parameterName, string itemType)
    {
        var dimensions = GetContentDimensions(range);
        RangeCommandValidation.ValidateDimensions(payload, dimensions.Rows, dimensions.Columns, parameterName, itemType);
    }

    private static Excel.Range ResolveCopyDestination(Excel.Range source, Excel.Range target, bool transpose)
    {
        var sourceSize = GetContentDimensions(source);
        if (transpose)
        {
            sourceSize = (sourceSize.Columns, sourceSize.Rows);
        }
        var targetSize = GetContentDimensions(target);
        int rows = targetSize.Rows;
        int columns = targetSize.Columns;
        if (rows == 1 && columns == 1)
        {
            rows = sourceSize.Rows;
            columns = sourceSize.Columns;
        }
        else if (rows % sourceSize.Rows != 0 || columns % sourceSize.Columns != 0)
        {
            throw new OperationFailureException(
                OperationFailureCategory.InvalidInput,
                "Cannot determine the copy destination: use a single-cell anchor or " +
                "a rectangle whose row and column counts are whole multiples of the source's paste dimensions.");
        }

        if ((long)target.Row + rows - 1 > 1_048_576 ||
            (long)target.Column + columns - 1 > 16_384)
        {
            throw new OperationFailureException(
                OperationFailureCategory.InvalidInput,
                "The expanded copy destination exceeds the worksheet boundaries.");
        }

        Excel.Range? destination = null;
        try
        {
            destination = target.Resize[rows, columns];
            if (RangeMergeDiscovery.GetMergeCellsState(source.MergeCells) != false ||
                RangeMergeDiscovery.GetMergeCellsState(destination.MergeCells) != false)
            {
                throw new OperationFailureException(
                    OperationFailureCategory.Conflict,
                    "Cannot determine the copy's complete destination when source or destination " +
                    "intersects merged cells. Use unmerged rectangular ranges; allow is only for authorized replacement.");
            }

            var result = destination;
            destination = null;
            return result;
        }
        finally
        {
            ComUtilities.Release(ref destination);
        }
    }

    private static void EnsureDestinationWritable(
        ExcelContext context, Excel.Range destination, OverwritePolicy overwritePolicy,
        CancellationToken cancellationToken, Func<int, int, bool>? writesCell = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (overwritePolicy == OverwritePolicy.Allow)
        {
            return;
        }

        var dimensions = GetContentDimensions(destination);
        int blockRows = Math.Max(1, InspectionBlockCells / dimensions.Columns);
        var conflicts = new List<string>();
        Excel.Worksheet? sheet = null;
        Excel.Range? cells = null;
        Excel.Range? anchor = null;
        try
        {
            sheet = destination.Worksheet;
            string sheetName = sheet.Name;
            cells = destination.Cells;
            anchor = cells[1, 1];
            int firstRow = destination.Row;
            int firstColumn = destination.Column;
            for (int rowOffset = 0; rowOffset < dimensions.Rows; rowOffset += blockRows)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int rows = Math.Min(blockRows, dimensions.Rows - rowOffset);
                Excel.Range? start = null;
                Excel.Range? block = null;
                try
                {
                    start = anchor.Offset[rowOffset, 0];
                    block = start.Resize[rows, dimensions.Columns];
                    object? values = block.Value2;
                    object? formulas = ReadFormulas(context, block);
                    for (int row = 0; row < rows; row++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        for (int column = 0; column < dimensions.Columns; column++)
                        {
                            object? value = GetInspectionCell(values, rows, dimensions.Columns, row, column);
                            object? formula = GetInspectionCell(formulas, rows, dimensions.Columns, row, column);
                            if (!RangeCommandValidation.IsOccupied(value, formula))
                            {
                                continue;
                            }
                            if (writesCell is not null && !writesCell(rowOffset + row, column))
                            {
                                continue;
                            }

                            if (conflicts.Count == ConflictExampleLimit)
                            {
                                ThrowOccupiedDestination(sheetName, conflicts, truncated: true);
                            }

                            conflicts.Add($"${GetColumnLetter(firstColumn + column)}${firstRow + rowOffset + row}");
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref block);
                    ComUtilities.Release(ref start);
                }
            }

            if (conflicts.Count > 0)
            {
                ThrowOccupiedDestination(sheetName, conflicts, truncated: false);
            }
        }
        finally
        {
            ComUtilities.Release(ref anchor);
            ComUtilities.Release(ref cells);
            ComUtilities.Release(ref sheet);
        }
    }

    private static object? GetInspectionCell(object? data, int rows, int columns, int row, int column)
    {
        if (data is object[,] cells && cells.GetLength(0) == rows && cells.GetLength(1) == columns)
        {
            return cells[row + cells.GetLowerBound(0), column + cells.GetLowerBound(1)];
        }

        if (rows == 1 && columns == 1 && data is not Array)
        {
            return data;
        }

        throw new InvalidOperationException(
            "Cannot inspect destination content: Excel returned an unexpected range shape. No write was attempted.");
    }

    private static void ThrowOccupiedDestination(string sheetName, List<string> conflicts, bool truncated)
    {
        RangeCommandValidation.ThrowOccupiedDestination(sheetName, conflicts, truncated);
    }
}
