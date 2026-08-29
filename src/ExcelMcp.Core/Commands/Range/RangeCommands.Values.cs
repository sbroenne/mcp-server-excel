using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;


namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Range value operations (get/set values as 2D arrays)
/// </summary>
public partial class RangeCommands
{
    /// <inheritdoc />
    public RangeValueResult GetValues(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        int rowOffset = 0,
        int? rowLimit = null,
        string? columns = null)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(rowOffset);
        if (rowLimit <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(rowLimit), rowLimit, "Row limit must be greater than zero.");
        }

        var selectedColumns = ParseSelectedColumns(columns);
        bool isScopedRead = rowOffset != 0 || rowLimit.HasValue || selectedColumns != null;
        var result = new RangeValueResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sheetName,
            RangeAddress = rangeAddress,
            RowOffset = rowOffset,
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
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    throw new InvalidOperationException(specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                }

                // Get actual address from Excel
                result.RangeAddress = range.Address;
                rows = range.Rows;
                rangeColumns = range.Columns;
                result.TotalRowCount = Convert.ToInt32(rows.Count);
                result.TotalColumnCount = Convert.ToInt32(rangeColumns.Count);

                if (rowOffset > result.TotalRowCount)
                {
                    throw new ArgumentOutOfRangeException(
                        nameof(rowOffset),
                        rowOffset,
                        $"Row offset must be between 0 and the source row count ({result.TotalRowCount}).");
                }

                if (isScopedRead)
                {
                    areas = range.Areas;
                    if (Convert.ToInt32(areas.Count) != 1)
                    {
                        throw new ArgumentException(
                            "Scoped reads do not support multi-area ranges. Read each area separately.",
                            nameof(rangeAddress));
                    }
                }

                int sourceStartColumn = Convert.ToInt32(range.Column);
                var resolvedColumns = ResolveSelectedColumns(
                    selectedColumns,
                    sourceStartColumn,
                    result.TotalColumnCount);
                int returnedColumnCount = resolvedColumns?.Count ?? result.TotalColumnCount;
                int remainingRows = result.TotalRowCount - rowOffset;
                int returnedRowCount = Math.Min(rowLimit ?? remainingRows, remainingRows);

                result.RowCount = returnedRowCount;
                result.ColumnCount = returnedColumnCount;
                result.HasMoreRows = rowOffset + returnedRowCount < result.TotalRowCount;
                result.NextRowOffset = result.HasMoreRows ? rowOffset + returnedRowCount : null;
                result.IsTruncated =
                    rowOffset > 0 ||
                    returnedRowCount < result.TotalRowCount ||
                    returnedColumnCount < result.TotalColumnCount;

                if (returnedRowCount == 0)
                {
                    result.Success = true;
                    return result;
                }

                if (!isScopedRead)
                {
                    PopulateValues(result.Values, range.Value2, returnedRowCount, returnedColumnCount);
                }
                else if (resolvedColumns == null)
                {
                    ReadRectangularSlice(
                        range,
                        rowOffset,
                        1,
                        returnedRowCount,
                        returnedColumnCount,
                        result.Values);
                }
                else
                {
                    for (int rowIndex = 0; rowIndex < returnedRowCount; rowIndex++)
                    {
                        result.Values.Add(new List<object?>(returnedColumnCount));
                    }

                    foreach (var selectedColumn in resolvedColumns)
                    {
                        ct.ThrowIfCancellationRequested();
                        var columnValues = new List<List<object?>>(returnedRowCount);
                        ReadRectangularSlice(
                            range,
                            rowOffset,
                            selectedColumn.RelativeIndex,
                            returnedRowCount,
                            1,
                            columnValues);

                        for (int rowIndex = 0; rowIndex < returnedRowCount; rowIndex++)
                        {
                            result.Values[rowIndex].Add(columnValues[rowIndex][0]);
                        }
                    }
                }

                result.Success = true;
                return result;
            }
            catch (System.Runtime.InteropServices.COMException comEx) when (comEx.HResult == unchecked((int)0x8007000E))
            {
                // E_OUTOFMEMORY - Excel's misleading error for sheet/range/session issues
                throw new InvalidOperationException($"Cannot read range '{rangeAddress}' on sheet '{sheetName}': {comEx.Message}", comEx);
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

    private static List<(string Name, int AbsoluteIndex)>? ParseSelectedColumns(string? columns)
    {
        if (string.IsNullOrWhiteSpace(columns))
        {
            return null;
        }

        var selectedColumns = new List<(string Name, int AbsoluteIndex)>();
        var seenColumns = new HashSet<int>();
        foreach (string rawColumn in columns.Split(','))
        {
            string columnName = rawColumn.Trim().ToUpperInvariant();
            if (columnName.Length > 3 ||
                !RangeHelpers.TryGetColumnIndex(columnName, out int absoluteIndex))
            {
                throw new ArgumentException(
                    $"Column '{rawColumn.Trim()}' is invalid. Use Excel column letters from A through XFD.",
                    nameof(columns));
            }

            if (!seenColumns.Add(absoluteIndex))
            {
                throw new ArgumentException($"Column '{columnName}' was specified more than once.", nameof(columns));
            }

            selectedColumns.Add((columnName, absoluteIndex));
        }

        return selectedColumns;
    }

    private static List<(string Name, int RelativeIndex)>? ResolveSelectedColumns(
        List<(string Name, int AbsoluteIndex)>? columns,
        int sourceStartColumn,
        int sourceColumnCount)
    {
        if (columns == null)
        {
            return null;
        }

        int sourceEndColumn = sourceStartColumn + sourceColumnCount - 1;
        var resolvedColumns = new List<(string Name, int RelativeIndex)>(columns.Count);
        foreach (var selectedColumn in columns)
        {
            if (selectedColumn.AbsoluteIndex < sourceStartColumn ||
                selectedColumn.AbsoluteIndex > sourceEndColumn)
            {
                throw new ArgumentException(
                    $"Column '{selectedColumn.Name}' is outside the resolved source range " +
                    $"({RangeHelpers.GetColumnLetter(sourceStartColumn)}:{RangeHelpers.GetColumnLetter(sourceEndColumn)}).",
                    nameof(columns));
            }

            resolvedColumns.Add((
                selectedColumn.Name,
                selectedColumn.AbsoluteIndex - sourceStartColumn + 1));
        }

        return resolvedColumns;
    }

    private static void ReadRectangularSlice(
        Excel.Range sourceRange,
        int rowOffset,
        int relativeColumnIndex,
        int rowCount,
        int columnCount,
        List<List<object?>> destination)
    {
        Excel.Range? cells = null;
        Excel.Range? firstCell = null;
        Excel.Range? slice = null;
        try
        {
            cells = sourceRange.Cells;
            firstCell = cells[rowOffset + 1, relativeColumnIndex];
            if (rowCount == 1 && columnCount == 1)
            {
                PopulateValues(destination, firstCell.Value2, rowCount, columnCount);
                return;
            }

            slice = firstCell.Resize[rowCount, columnCount];
            PopulateValues(destination, slice.Value2, rowCount, columnCount);
        }
        finally
        {
            ComUtilities.Release(ref slice);
            ComUtilities.Release(ref firstCell);
            ComUtilities.Release(ref cells);
        }
    }

    private static void PopulateValues(
        List<List<object?>> destination,
        object? valueOrArray,
        int rowCount,
        int columnCount)
    {
        if (valueOrArray is object[,] values)
        {
            for (int rowIndex = 1; rowIndex <= rowCount; rowIndex++)
            {
                var row = new List<object?>(columnCount);
                for (int columnIndex = 1; columnIndex <= columnCount; columnIndex++)
                {
                    row.Add(values[rowIndex, columnIndex]);
                }

                destination.Add(row);
            }

            return;
        }

        destination.Add([valueOrArray]);
    }

    /// <inheritdoc />
    public OperationResult SetValues(IExcelBatch batch, string sheetName, string rangeAddress, List<List<object?>>? values = null, string? valuesFile = null)
    {
        // Resolve values from inline parameter or file
        var resolvedValues = ParameterTransforms.ResolveValuesOrFile(values, valuesFile);

        // SMART FORMULA DETECTION: Check if any value starts with "=" and auto-route to SetFormulas
        bool hasFormulas = DetectFormulas(resolvedValues, out var detectedFormulas);
        if (hasFormulas)
        {
            // Detected formulas - convert to proper formula format and use SetFormulas
            var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "set-values" };

            // Call SetFormulas internally to apply detected formulas
            var formulaResult = SetFormulas(batch, sheetName, rangeAddress, detectedFormulas);

            // Copy result data and add detection message
            result.Success = formulaResult.Success;
            result.ErrorMessage = formulaResult.ErrorMessage;
            if (result.Success && string.IsNullOrEmpty(result.Message))
            {
                result.Message = $"Formula detected: {detectedFormulas.Sum(row => row.Count(f => !string.IsNullOrEmpty(f)))} formula(s) applied via set-formulas";
            }
            return result;
        }

        var setResult = new OperationResult { FilePath = batch.WorkbookPath, Action = "set-values" };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            int originalCalculation = -1;
            bool calculationChanged = false;

            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    throw new InvalidOperationException(specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                }

                // Calculation suppressed here (not in ExcelWriteGuard) because Data Model ops need it enabled
                originalCalculation = (int)ctx.App.Calculation;
                if (originalCalculation != -4135) // xlCalculationManual
                {
                    ctx.App.Calculation = (Excel.XlCalculation)(-4135);
                    calculationChanged = true;
                }

                // Convert List<List<object?>> to 2D array
                // Excel COM requires 1-based arrays for multi-cell ranges
                int rows = resolvedValues.Count;
                int cols = resolvedValues.Count > 0 ? resolvedValues[0].Count : 0;

                ValidateRectangularRowWidths(resolvedValues, Convert.ToInt32(range.Columns.Count), nameof(values), "Value");

                if (rows > 0 && cols > 0)
                {
                    // Create 1-based array for Excel COM compatibility
                    object[,] arrayValues = (object[,])Array.CreateInstance(typeof(object), [rows, cols], [1, 1]);

                    for (int r = 1; r <= rows; r++)
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            // Convert JsonElement to proper C# type for COM interop
                            // MCP framework deserializes JSON to JsonElement, not primitives
                            arrayValues[r, c] = RangeHelpers.ConvertToCellValue(resolvedValues[r - 1][c - 1]);
                        }
                    }

                    range.Value2 = arrayValues;
                }

                setResult.Success = true;
                return setResult;
            }
            catch (System.Runtime.InteropServices.COMException comEx) when (comEx.HResult == unchecked((int)0x8007000E))
            {
                // E_OUTOFMEMORY - Excel's misleading error for sheet/range/session issues
                throw new InvalidOperationException($"Cannot write to range '{rangeAddress}' on sheet '{sheetName}': {comEx.Message}", comEx);
            }
            finally
            {
                if (calculationChanged && originalCalculation != -1)
                {
                    try
                    {
                        ctx.App.Calculation = (Excel.XlCalculation)originalCalculation;
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        // Ignore errors restoring calculation mode
                    }
                }
                ComUtilities.Release(ref range);
            }
        });
    }

    /// <summary>
    /// Detects formulas in value array (strings starting with =)
    /// Returns true if any formulas detected, outputs formula array
    /// </summary>
    private static bool DetectFormulas(List<List<object?>> values, out List<List<string>> detectedFormulas)
    {
        detectedFormulas = new List<List<string>>();
        bool hasFormulas = false;

        foreach (var row in values)
        {
            var formulaRow = new List<string>();
            foreach (var value in row)
            {
                string str = value?.ToString() ?? string.Empty;

                // Detect formula (starts with = but not escaped with ')
                if (str.StartsWith('=') && !str.StartsWith("'=", StringComparison.Ordinal))
                {
                    formulaRow.Add(str);
                    hasFormulas = true;
                }
                else
                {
                    // Not a formula - empty string in formula array
                    formulaRow.Add(string.Empty);
                }
            }
            detectedFormulas.Add(formulaRow);
        }

        return hasFormulas;
    }

    /// <summary>
    /// Validates that every row in a 2D payload matches the target range width before COM indexing.
    /// </summary>
    private static void ValidateRectangularRowWidths<T>(List<List<T>> rows, int expectedColumnCount, string parameterName, string itemType)
    {
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            if (rows[rowIndex].Count != expectedColumnCount)
            {
                throw new ArgumentException(
                    $"{itemType} array row {rowIndex + 1} column count ({rows[rowIndex].Count}) doesn't match range column count ({expectedColumnCount})",
                    parameterName);
            }
        }
    }
}
