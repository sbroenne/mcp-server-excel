using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Bounded range sampling, summaries, and sparse formula error diagnostics.
/// </summary>
public partial class RangeCommands
{
    private const int MaxSampleRowsPerBoundary = 100;
    private const int MaxSampleCells = 1000;
    private const int MaxSummaryColumns = 256;
    private const int MaxFormulaErrors = 1000;
    private const int MaxSpecialCellsChunkSize = 4096;
    private const int NoCellsFoundHResult = unchecked((int)0x800A03EC);

    /// <inheritdoc />
    public RangeSampleResult SampleValues(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        int firstRowCount = 5,
        int lastRowCount = 5,
        string? columns = null)
    {
        ValidateSampleCount(firstRowCount, nameof(firstRowCount));
        ValidateSampleCount(lastRowCount, nameof(lastRowCount));
        if (firstRowCount == 0 && lastRowCount == 0)
        {
            throw new ArgumentException(
                "At least one of firstRowCount or lastRowCount must be greater than zero.",
                nameof(firstRowCount));
        }

        var selectedColumns = ParseSelectedColumns(columns);
        var result = new RangeSampleResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sheetName,
            RangeAddress = rangeAddress,
            FirstRowCount = firstRowCount,
            LastRowCount = lastRowCount,
            SelectedColumns = selectedColumns?.Select(column => column.Name).ToList()
        };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Range? range = null;
            Excel.Range? rows = null;
            Excel.Range? rangeColumns = null;
            Excel.Areas? areas = null;
            try
            {
                range = ResolveSingleAreaRange(ctx.Book, sheetName, rangeAddress, "Sampling", out areas);
                result.RangeAddress = range.Address;
                rows = range.Rows;
                rangeColumns = range.Columns;
                result.TotalRowCount = Convert.ToInt32(rows.Count);
                result.TotalColumnCount = Convert.ToInt32(rangeColumns.Count);

                int sourceStartColumn = Convert.ToInt32(range.Column);
                int sourceStartRow = Convert.ToInt32(range.Row);
                var resolvedColumns = ResolveSelectedColumns(
                    selectedColumns,
                    sourceStartColumn,
                    result.TotalColumnCount);
                result.ColumnCount = resolvedColumns?.Count ?? result.TotalColumnCount;

                int firstCount = Math.Min(firstRowCount, result.TotalRowCount);
                int lastStart = Math.Max(firstCount, result.TotalRowCount - lastRowCount);
                int returnedRowCount = firstCount + result.TotalRowCount - lastStart;
                if ((long)returnedRowCount * result.ColumnCount > MaxSampleCells)
                {
                    throw new ArgumentException(
                        $"The sample would return more than {MaxSampleCells} cells. Reduce the row counts or select fewer columns.",
                        nameof(columns));
                }

                var rowOffsets = Enumerable.Range(0, firstCount)
                    .Concat(Enumerable.Range(lastStart, result.TotalRowCount - lastStart));
                foreach (int rowOffset in rowOffsets)
                {
                    ct.ThrowIfCancellationRequested();
                    var values = ReadSampleRow(
                        range,
                        rowOffset,
                        resolvedColumns,
                        result.TotalColumnCount);
                    int rowNumber = sourceStartRow + rowOffset;
                    result.Rows.Add(new RangeSampleRow
                    {
                        RowOffset = rowOffset,
                        RowNumber = rowNumber,
                        RangeAddress = BuildSampleRowAddress(
                            rowNumber,
                            sourceStartColumn,
                            result.TotalColumnCount,
                            resolvedColumns),
                        Values = values
                    });
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref areas);
                ComUtilities.Release(ref rangeColumns);
                ComUtilities.Release(ref rows);
                ComUtilities.Release(ref range);
            }
        });
    }

    /// <inheritdoc />
    public RangeSummaryResult SummarizeValues(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        string? columns = null)
    {
        var selectedColumns = ParseSelectedColumns(columns);
        if (selectedColumns?.Count > MaxSummaryColumns)
        {
            throw new ArgumentException(
                $"At most {MaxSummaryColumns} columns can be summarized.",
                nameof(columns));
        }

        var result = new RangeSummaryResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sheetName,
            RangeAddress = rangeAddress,
            SelectedColumns = selectedColumns?.Select(column => column.Name).ToList()
        };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Range? range = null;
            Excel.Range? rows = null;
            Excel.Range? rangeColumns = null;
            Excel.Areas? areas = null;
            Excel.WorksheetFunction? worksheetFunction = null;
            try
            {
                range = ResolveSingleAreaRange(ctx.Book, sheetName, rangeAddress, "Summaries", out areas);
                result.RangeAddress = range.Address;
                rows = range.Rows;
                rangeColumns = range.Columns;
                result.TotalRowCount = Convert.ToInt32(rows.Count);
                result.TotalColumnCount = Convert.ToInt32(rangeColumns.Count);

                int sourceStartColumn = Convert.ToInt32(range.Column);
                var resolvedColumns = ResolveSelectedColumns(
                    selectedColumns,
                    sourceStartColumn,
                    result.TotalColumnCount);
                int returnedColumnCount = resolvedColumns?.Count ?? result.TotalColumnCount;
                if (returnedColumnCount > MaxSummaryColumns)
                {
                    throw new ArgumentException(
                        $"The source has {returnedColumnCount} columns. Select at most {MaxSummaryColumns} columns.",
                        nameof(columns));
                }

                worksheetFunction = ctx.App.WorksheetFunction;
                for (int index = 0; index < returnedColumnCount; index++)
                {
                    ct.ThrowIfCancellationRequested();
                    int relativeColumn = resolvedColumns?[index].RelativeIndex ?? index + 1;
                    int absoluteColumn = sourceStartColumn + relativeColumn - 1;
                    Excel.Range? cells = null;
                    Excel.Range? firstCell = null;
                    Excel.Range? columnRange = null;
                    try
                    {
                        cells = range.Cells;
                        firstCell = cells[1, relativeColumn];
                        columnRange = firstCell.Resize[result.TotalRowCount, 1];
                        result.Columns.Add(SummarizeColumn(
                            columnRange,
                            worksheetFunction,
                            absoluteColumn,
                            ct));
                    }
                    finally
                    {
                        ComUtilities.Release(ref columnRange);
                        ComUtilities.Release(ref firstCell);
                        ComUtilities.Release(ref cells);
                    }
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref worksheetFunction);
                ComUtilities.Release(ref areas);
                ComUtilities.Release(ref rangeColumns);
                ComUtilities.Release(ref rows);
                ComUtilities.Release(ref range);
            }
        });
    }

    /// <inheritdoc />
    public RangeFormulaErrorResult GetFormulaErrors(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        int maxErrors = 100)
    {
        if (maxErrors is < 1 or > MaxFormulaErrors)
        {
            throw new ArgumentOutOfRangeException(
                nameof(maxErrors),
                maxErrors,
                $"Maximum errors must be between 1 and {MaxFormulaErrors}.");
        }

        var result = new RangeFormulaErrorResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sheetName,
            RangeAddress = rangeAddress,
            MaxErrors = maxErrors
        };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Range? range = null;
            Excel.Areas? sourceAreas = null;
            var orderedSourceAreas = new List<(int Index, int Row, int Column)>();
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    throw new InvalidOperationException(
                        specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                }

                result.RangeAddress = range.Address;
                sourceAreas = range.Areas;
                int sourceAreaCount = Convert.ToInt32(sourceAreas.Count);
                for (int sourceAreaIndex = 1; sourceAreaIndex <= sourceAreaCount; sourceAreaIndex++)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.Range? sourceArea = null;
                    try
                    {
                        sourceArea = sourceAreas[sourceAreaIndex];
                        orderedSourceAreas.Add((
                            sourceAreaIndex,
                            Convert.ToInt32(sourceArea.Row),
                            Convert.ToInt32(sourceArea.Column)));
                        result.TotalErrorCount += CountFormulaErrors(sourceArea, ct);
                    }
                    finally
                    {
                        ComUtilities.Release(ref sourceArea);
                    }
                }

                foreach (var sourceAreaOrder in orderedSourceAreas
                             .OrderBy(area => area.Row)
                             .ThenBy(area => area.Column))
                {
                    if (result.Errors.Count >= maxErrors)
                    {
                        break;
                    }

                    ct.ThrowIfCancellationRequested();
                    Excel.Range? sourceArea = null;
                    try
                    {
                        sourceArea = sourceAreas[sourceAreaOrder.Index];
                        AppendFormulaErrors(sourceArea, result.Errors, maxErrors, ct);
                    }
                    finally
                    {
                        ComUtilities.Release(ref sourceArea);
                    }
                }

                result.ReturnedErrorCount = result.Errors.Count;
                result.IsTruncated = result.ReturnedErrorCount < result.TotalErrorCount;
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref sourceAreas);
                ComUtilities.Release(ref range);
            }
        });
    }

    private static void ValidateSampleCount(int count, string parameterName)
    {
        if (count is < 0 or > MaxSampleRowsPerBoundary)
        {
            throw new ArgumentOutOfRangeException(
                parameterName,
                count,
                $"Sample row counts must be between 0 and {MaxSampleRowsPerBoundary}.");
        }
    }

    private static Excel.Range ResolveSingleAreaRange(
        Excel.Workbook workbook,
        string sheetName,
        string rangeAddress,
        string operation,
        out Excel.Areas areas)
    {
        Excel.Range? range = RangeHelpers.ResolveRange(
            workbook,
            sheetName,
            rangeAddress,
            out string? specificError);
        if (range == null)
        {
            throw new InvalidOperationException(
                specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
        }

        try
        {
            areas = range.Areas;
            if (Convert.ToInt32(areas.Count) != 1)
            {
                throw new ArgumentException(
                    $"{operation} do not support multi-area ranges. Read each area separately.",
                    nameof(rangeAddress));
            }

            return range;
        }
        catch
        {
            ComUtilities.Release(ref range);
            throw;
        }
    }

    private static List<object?> ReadSampleRow(
        Excel.Range range,
        int rowOffset,
        List<(string Name, int RelativeIndex)>? resolvedColumns,
        int totalColumnCount)
    {
        if (resolvedColumns == null)
        {
            var values = new List<List<object?>>(1);
            ReadRectangularSlice(range, rowOffset, 1, 1, totalColumnCount, values);
            return values[0];
        }

        var selectedValues = new List<object?>(resolvedColumns.Count);
        foreach (var selectedColumn in resolvedColumns)
        {
            var values = new List<List<object?>>(1);
            ReadRectangularSlice(range, rowOffset, selectedColumn.RelativeIndex, 1, 1, values);
            selectedValues.Add(values[0][0]);
        }

        return selectedValues;
    }

    private static string BuildSampleRowAddress(
        int rowNumber,
        int sourceStartColumn,
        int totalColumnCount,
        List<(string Name, int RelativeIndex)>? resolvedColumns)
    {
        if (resolvedColumns != null)
        {
            return string.Join(",", resolvedColumns.Select(column => $"${column.Name}${rowNumber}"));
        }

        string firstColumn = RangeHelpers.GetColumnLetter(sourceStartColumn);
        string lastColumn = RangeHelpers.GetColumnLetter(sourceStartColumn + totalColumnCount - 1);
        return totalColumnCount == 1
            ? $"${firstColumn}${rowNumber}"
            : $"${firstColumn}${rowNumber}:${lastColumn}${rowNumber}";
    }

    private static RangeColumnSummary SummarizeColumn(
        Excel.Range columnRange,
        Excel.WorksheetFunction worksheetFunction,
        int absoluteColumn,
        CancellationToken cancellationToken)
    {
        long cellCount = Convert.ToInt64(columnRange.CountLarge);
        var summary = new RangeColumnSummary
        {
            Column = RangeHelpers.GetColumnLetter(absoluteColumn),
            ColumnNumber = absoluteColumn,
            RangeAddress = columnRange.Address,
            CellCount = cellCount
        };

        var numeric = new NumericSummaryAccumulator();
        ForEachRangeChunk(columnRange, cancellationToken, chunk =>
        {
            if (Convert.ToInt64(chunk.CountLarge) == 1)
            {
                AccumulateSingleCellSummary(chunk.Value2, summary, numeric);
            }
            else
            {
                summary.NumericCount += AccumulateSpecialCells(
                    chunk,
                    Excel.XlSpecialCellsValue.xlNumbers,
                    worksheetFunction,
                    numeric);
                summary.TextCount += CountSpecialCells(
                    chunk,
                    Excel.XlSpecialCellsValue.xlTextValues);
                summary.LogicalCount += CountSpecialCells(
                    chunk,
                    Excel.XlSpecialCellsValue.xlLogical);
                summary.ErrorCount += CountSpecialCells(
                    chunk,
                    Excel.XlSpecialCellsValue.xlErrors);
            }

            return true;
        });
        summary.BlankCount = cellCount -
            summary.NumericCount -
            summary.TextCount -
            summary.LogicalCount -
            summary.ErrorCount;

        if (summary.NumericCount > 0)
        {
            summary.Sum = numeric.Sum;
            summary.Average = numeric.Sum / summary.NumericCount;
            summary.Minimum = numeric.Minimum;
            summary.Maximum = numeric.Maximum;
        }

        return summary;
    }

    private static void AccumulateSingleCellSummary(
        object? value,
        RangeColumnSummary summary,
        NumericSummaryAccumulator numeric)
    {
        if (value == null)
        {
            return;
        }

        if (ExcelErrorMapper.TryGetErrorCode(value, out _))
        {
            summary.ErrorCount++;
        }
        else if (value is bool)
        {
            summary.LogicalCount++;
        }
        else if (value is string)
        {
            summary.TextCount++;
        }
        else if (value is byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal)
        {
            double number = Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture);
            summary.NumericCount++;
            numeric.Add(number, number, number);
        }
        else
        {
            summary.TextCount++;
        }
    }

    private static long CountSpecialCells(
        Excel.Range source,
        Excel.XlSpecialCellsValue valueType)
    {
        long count = 0;
        foreach (Excel.XlCellType cellType in new[]
                 {
                     Excel.XlCellType.xlCellTypeConstants,
                     Excel.XlCellType.xlCellTypeFormulas
                 })
        {
            Excel.Range? cells = null;
            try
            {
                cells = TryGetSpecialCells(source, cellType, valueType);
                if (cells != null)
                {
                    count += Convert.ToInt64(cells.CountLarge);
                }
            }
            finally
            {
                ComUtilities.Release(ref cells);
            }
        }

        return count;
    }

    private static long AccumulateSpecialCells(
        Excel.Range source,
        Excel.XlSpecialCellsValue valueType,
        Excel.WorksheetFunction worksheetFunction,
        NumericSummaryAccumulator accumulator)
    {
        long count = 0;
        foreach (Excel.XlCellType cellType in new[]
                 {
                     Excel.XlCellType.xlCellTypeConstants,
                     Excel.XlCellType.xlCellTypeFormulas
                 })
        {
            Excel.Range? cells = null;
            try
            {
                cells = TryGetSpecialCells(source, cellType, valueType);
                if (cells == null)
                {
                    continue;
                }

                long subsetCount = Convert.ToInt64(cells.CountLarge);
                count += subsetCount;
                accumulator.Add(
                    Convert.ToDouble(worksheetFunction.Sum(cells)),
                    Convert.ToDouble(worksheetFunction.Min(cells)),
                    Convert.ToDouble(worksheetFunction.Max(cells)));
            }
            finally
            {
                ComUtilities.Release(ref cells);
            }
        }

        return count;
    }

    private static Excel.Range? TryGetSpecialCells(
        Excel.Range source,
        Excel.XlCellType cellType,
        Excel.XlSpecialCellsValue valueType)
    {
        try
        {
            return source.SpecialCells(cellType, valueType);
        }
        catch (COMException exception) when (exception.HResult == NoCellsFoundHResult)
        {
            return null;
        }
    }

    private static long CountFormulaErrors(
        Excel.Range sourceArea,
        CancellationToken cancellationToken)
    {
        long count = 0;
        ForEachRangeChunk(sourceArea, cancellationToken, chunk =>
        {
            if (Convert.ToInt64(chunk.CountLarge) == 1)
            {
                object? value = chunk.Value2;
                string formula = chunk.Formula2?.ToString() ?? string.Empty;
                if (formula.StartsWith('=') &&
                    ExcelErrorMapper.TryGetErrorCode(value, out _))
                {
                    count++;
                }
            }
            else
            {
                Excel.Range? formulaErrors = null;
                try
                {
                    formulaErrors = TryGetSpecialCells(
                        chunk,
                        Excel.XlCellType.xlCellTypeFormulas,
                        Excel.XlSpecialCellsValue.xlErrors);
                    if (formulaErrors != null)
                    {
                        count += Convert.ToInt64(formulaErrors.CountLarge);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref formulaErrors);
                }
            }

            return true;
        });
        return count;
    }

    private static void AppendFormulaErrors(
        Excel.Range sourceArea,
        List<RangeCellError> destination,
        int maxErrors,
        CancellationToken cancellationToken)
    {
        ForEachRangeChunk(sourceArea, cancellationToken, chunk =>
        {
            if (Convert.ToInt64(chunk.CountLarge) == 1)
            {
                if (TryCreateFormulaError(chunk, out var error))
                {
                    destination.Add(error);
                }
            }
            else
            {
                Excel.Range? formulaErrors = null;
                Excel.Areas? errorAreas = null;
                try
                {
                    formulaErrors = TryGetSpecialCells(
                        chunk,
                        Excel.XlCellType.xlCellTypeFormulas,
                        Excel.XlSpecialCellsValue.xlErrors);
                    if (formulaErrors == null)
                    {
                        return destination.Count < maxErrors;
                    }

                    errorAreas = formulaErrors.Areas;
                    int errorAreaCount = Convert.ToInt32(errorAreas.Count);
                    for (int areaIndex = 1;
                         areaIndex <= errorAreaCount && destination.Count < maxErrors;
                         areaIndex++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        Excel.Range? errorArea = null;
                        Excel.Range? cells = null;
                        try
                        {
                            errorArea = errorAreas[areaIndex];
                            cells = errorArea.Cells;
                            long cellCount = Convert.ToInt64(cells.CountLarge);
                            long cellsToRead = Math.Min(cellCount, maxErrors - destination.Count);
                            for (long cellIndex = 1; cellIndex <= cellsToRead; cellIndex++)
                            {
                                cancellationToken.ThrowIfCancellationRequested();
                                Excel.Range? cell = null;
                                try
                                {
                                    cell = cells[cellIndex];
                                    if (TryCreateFormulaError(cell, out var error))
                                    {
                                        destination.Add(error);
                                    }
                                }
                                finally
                                {
                                    ComUtilities.Release(ref cell);
                                }
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref cells);
                            ComUtilities.Release(ref errorArea);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref errorAreas);
                    ComUtilities.Release(ref formulaErrors);
                }
            }

            return destination.Count < maxErrors;
        });
    }

    private static bool TryCreateFormulaError(
        Excel.Range cell,
        out RangeCellError error)
    {
        object? value = cell.Value2;
        string formula = cell.Formula2?.ToString() ?? string.Empty;
        if (!formula.StartsWith('=') ||
            !ExcelErrorMapper.TryGetErrorCode(value, out int errorCode))
        {
            error = null!;
            return false;
        }

        error = new RangeCellError
        {
            CellAddress = cell.Address[RowAbsolute: false, ColumnAbsolute: false],
            Row = Convert.ToInt32(cell.Row),
            Column = Convert.ToInt32(cell.Column),
            Formula = formula,
            CurrentValue = value,
            ErrorCode = errorCode,
            ErrorMessage = ExcelErrorMapper.GetMessage(errorCode),
            Suggestion = ExcelErrorMapper.GetSuggestion(errorCode)
        };
        return true;
    }

    private static void ForEachRangeChunk(
        Excel.Range source,
        CancellationToken cancellationToken,
        Func<Excel.Range, bool> visitor)
    {
        Excel.Range? rows = null;
        Excel.Range? columns = null;
        Excel.Range? cells = null;
        try
        {
            rows = source.Rows;
            columns = source.Columns;
            cells = source.Cells;
            int totalRows = Convert.ToInt32(rows.Count);
            int totalColumns = Convert.ToInt32(columns.Count);
            int columnChunkSize = Math.Min(totalColumns, MaxSpecialCellsChunkSize);
            int rowChunkSize = Math.Max(1, MaxSpecialCellsChunkSize / columnChunkSize);

            for (int rowStart = 1; rowStart <= totalRows; rowStart += rowChunkSize)
            {
                int rowsInChunk = Math.Min(rowChunkSize, totalRows - rowStart + 1);
                for (int columnStart = 1;
                     columnStart <= totalColumns;
                     columnStart += columnChunkSize)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    int columnsInChunk = Math.Min(
                        columnChunkSize,
                        totalColumns - columnStart + 1);
                    Excel.Range? firstCell = null;
                    Excel.Range? chunk = null;
                    try
                    {
                        firstCell = cells[rowStart, columnStart];
                        if (rowsInChunk == 1 && columnsInChunk == 1)
                        {
                            chunk = firstCell;
                            firstCell = null;
                        }
                        else
                        {
                            chunk = firstCell.Resize[rowsInChunk, columnsInChunk];
                        }

                        if (!visitor(chunk))
                        {
                            return;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref chunk);
                        ComUtilities.Release(ref firstCell);
                    }
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref cells);
            ComUtilities.Release(ref columns);
            ComUtilities.Release(ref rows);
        }
    }

    private sealed class NumericSummaryAccumulator
    {
        internal double Sum { get; private set; }

        internal double? Minimum { get; private set; }

        internal double? Maximum { get; private set; }

        internal void Add(double sum, double minimum, double maximum)
        {
            Sum += sum;
            Minimum = Minimum.HasValue ? Math.Min(Minimum.Value, minimum) : minimum;
            Maximum = Maximum.HasValue ? Math.Max(Maximum.Value, maximum) : maximum;
        }
    }

}
