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
    public RangeValueResult GetValues(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        var result = new RangeValueResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sheetName,
            RangeAddress = rangeAddress
        };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    throw new InvalidOperationException(specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                }

                // Get actual address from Excel
                result.RangeAddress = range.Address;
                // Get values as 2D array - handle single cell case
                object valueOrArray = range.Value2;
                object? formulaOrArray = null;
                bool formulasRead = false;
                int startRow = Convert.ToInt32(range.Row);
                int startColumn = Convert.ToInt32(range.Column);

                if (valueOrArray is object[,] values)
                {
                    // Multi-cell range - process as 2D array
                    result.RowCount = values.GetLength(0);
                    result.ColumnCount = values.GetLength(1);

                    for (int r = 1; r <= result.RowCount; r++)
                    {
                        var row = new List<object?>();
                        for (int c = 1; c <= result.ColumnCount; c++)
                        {
                            object? cellValue = values[r, c];
                            if (!ExcelErrorMapper.TryGet(cellValue, out int errorCode, out var error))
                            {
                                row.Add(cellValue);
                                continue;
                            }

                            if (!formulasRead)
                            {
                                formulaOrArray = ReadFormulas(ctx, (Excel.Range)range);
                                formulasRead = true;
                            }

                            string formula = formulaOrArray is object[,] formulas
                                ? GetReturnedFormula(formulas[r, c])
                                : string.Empty;
                            row.Add(ConvertMappedErrorForRead(
                                cellValue,
                                formula,
                                startRow + r - 1,
                                startColumn + c - 1,
                                result.CellErrors,
                                errorCode,
                                error));
                        }
                        result.Values.Add(row);
                    }
                }
                else
                {
                    // Single cell - wrap value in 1x1 array
                    result.RowCount = 1;
                    result.ColumnCount = 1;
                    if (ExcelErrorMapper.TryGet(valueOrArray, out int errorCode, out var error))
                    {
                        formulaOrArray = ReadFormulas(ctx, (Excel.Range)range);
                        result.Values.Add([
                            ConvertMappedErrorForRead(
                                valueOrArray,
                                GetReturnedFormula(formulaOrArray),
                                startRow,
                                startColumn,
                                result.CellErrors,
                                errorCode,
                                error)
                        ]);
                    }
                    else
                    {
                        result.Values.Add([valueOrArray]);
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
                ComUtilities.Release(ref range);
            }
        });
    }

    private static string GetReturnedFormula(object? formula)
    {
        string text = formula?.ToString() ?? string.Empty;
        return text.StartsWith('=') ? text : string.Empty;
    }

    /// <inheritdoc />
    public OperationResult SetValues(IExcelBatch batch, string sheetName, string rangeAddress, List<List<object?>>? values = null, string? valuesFile = null, OverwritePolicy overwritePolicy = OverwritePolicy.RejectNonempty)
    {
        ValidateOverwritePolicy(overwritePolicy);
        // Resolve values from inline parameter or file
        var resolvedValues = ParameterTransforms.ResolveValuesOrFile(values, valuesFile);

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

                ValidateMergedCellsForWrite((Excel.Range)range, rangeAddress, ct);
                ValidateContentWriteDimensions((Excel.Range)range, resolvedValues, nameof(values), "Value");
                EnsureDestinationWritable(ctx, (Excel.Range)range, overwritePolicy, ct);

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

                if (rows > 0 && cols > 0)
                {
                    // Create 1-based array for Excel COM compatibility
                    object[,] arrayValues = (object[,])Array.CreateInstance(typeof(object), [rows, cols], [1, 1]);
                    int formulaCount = 0;

                    for (int r = 1; r <= rows; r++)
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            // Convert JsonElement to proper C# type for COM interop
                            // MCP framework deserializes JSON to JsonElement, not primitives
                            object cellValue = RangeHelpers.ConvertToCellValue(resolvedValues[r - 1][c - 1]);
                            if (cellValue is string text && text.StartsWith('='))
                            {
                                formulaCount++;
                            }

                            arrayValues[r, c] = cellValue;
                        }
                    }

                    if (formulaCount > 0)
                    {
                        // Formula/Formula2 store constants exactly like Value2, so one mixed array
                        // writes the formulas without blanking the other cells (issue #1065).
                        if (ctx.Capabilities.SupportsFormula2)
                            ((Excel.Range)range).Formula2 = arrayValues;
                        else
                            ((Excel.Range)range).Formula = arrayValues;

                        setResult.Message = $"Formula detected: {formulaCount} formula(s) applied with set-formulas semantics; other cells kept as values";
                    }
                    else
                    {
                        range.Value2 = arrayValues;
                    }
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

    private static void ValidateMergedCellsForWrite(
        Excel.Range range,
        string requestedRangeAddress,
        CancellationToken cancellationToken)
    {
        object? mergeCells = range.MergeCells;
        bool? isMergedState = RangeMergeDiscovery.GetMergeCellsState(mergeCells);
        if (isMergedState == false)
        {
            return;
        }

        if (isMergedState == true && Convert.ToInt64(range.CountLarge) == 1)
        {
            Excel.Range? mergeArea = null;
            try
            {
                mergeArea = range.MergeArea;
                if (range.Row == mergeArea.Row && range.Column == mergeArea.Column)
                {
                    return;
                }

                ThrowMergedCellWriteError(
                    requestedRangeAddress,
                    [mergeArea.Address[true, true]]);
            }
            finally
            {
                ComUtilities.Release(ref mergeArea);
            }
        }

        var mergedRanges = RangeMergeDiscovery.CollectMergedRanges(
            range,
            isMergedState,
            cancellationToken);
        if (mergedRanges.Count > 0)
        {
            ThrowMergedCellWriteError(requestedRangeAddress, mergedRanges);
        }
    }

    private static void ThrowMergedCellWriteError(
        string requestedRangeAddress,
        List<string> mergedRanges)
    {
        string rangeLabel = mergedRanges.Count == 1 ? "Merged range" : "Merged ranges";
        throw new OperationFailureException(
            OperationFailureCategory.Conflict,
            $"Cannot write to range '{requestedRangeAddress}' because the write intersects merged cells. " +
            $"{rangeLabel}: {string.Join(", ", mergedRanges)}. " +
            "Write only to each merged range's top-left cell, or unmerge the affected range before writing.");
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
