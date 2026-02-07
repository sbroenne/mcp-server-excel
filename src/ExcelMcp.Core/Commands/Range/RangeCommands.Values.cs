using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;


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
                            row.Add(values[r, c]);
                        }
                        result.Values.Add(row);
                    }
                }
                else
                {
                    // Single cell - wrap value in 1x1 array
                    result.RowCount = 1;
                    result.ColumnCount = 1;
                    result.Values.Add([valueOrArray]);
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

    /// <inheritdoc />
    public OperationResult SetValues(IExcelBatch batch, string sheetName, string rangeAddress, List<List<object?>> values)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "set-values" };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            int originalCalculation = -1; // xlCalculationAutomatic = -4105, xlCalculationManual = -4135
            bool calculationChanged = false;

            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    throw new InvalidOperationException(specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                }

                // CRITICAL: Temporarily disable automatic calculation to prevent Excel from
                // hanging when changed values trigger dependent formulas that reference Data Model/DAX.
                // Without this, setting values can block the COM interface during recalculation.
                originalCalculation = ctx.App.Calculation;
                if (originalCalculation != -4135) // xlCalculationManual
                {
                    ctx.App.Calculation = -4135; // xlCalculationManual
                    calculationChanged = true;
                }

                // Convert List<List<object?>> to 2D array
                // Excel COM requires 1-based arrays for multi-cell ranges
                int rows = values.Count;
                int cols = values.Count > 0 ? values[0].Count : 0;

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
                            arrayValues[r, c] = RangeHelpers.ConvertToCellValue(values[r - 1][c - 1]);
                        }
                    }

                    range.Value2 = arrayValues;
                }

                result.Success = true;
                return result;
            }
            catch (System.Runtime.InteropServices.COMException comEx) when (comEx.HResult == unchecked((int)0x8007000E))
            {
                // E_OUTOFMEMORY - Excel's misleading error for sheet/range/session issues
                throw new InvalidOperationException($"Cannot write to range '{rangeAddress}' on sheet '{sheetName}': {comEx.Message}", comEx);
            }
            finally
            {
                // Restore original calculation mode
                if (calculationChanged && originalCalculation != -1)
                {
                    try
                    {
                        ctx.App.Calculation = originalCalculation;
                    }
                    catch
                    {
                        // Ignore errors restoring calculation mode - not critical
                    }
                }
                ComUtilities.Release(ref range);
            }
        });
    }
}



