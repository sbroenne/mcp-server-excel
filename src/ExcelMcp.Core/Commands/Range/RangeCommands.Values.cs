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
    public async Task<RangeValueResult> GetValuesAsync(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        var result = new RangeValueResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sheetName,
            RangeAddress = rangeAddress
        };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    result.Success = false;
                    result.ErrorMessage = specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress);
                    return result;
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
                result.Success = false;
                result.ErrorMessage = $"Cannot read range '{rangeAddress}' on sheet '{sheetName}': {comEx.Message}";
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref range);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> SetValuesAsync(IExcelBatch batch, string sheetName, string rangeAddress, List<List<object?>> values)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "set-values" };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    result.Success = false;
                    result.ErrorMessage = specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress);
                    return result;
                }

                // Convert List<List<object?>> to 2D array (0-based for Excel COM)
                int rows = values.Count;
                int cols = values.Count > 0 ? values[0].Count : 0;

                if (rows > 0 && cols > 0)
                {
                    object[,] arrayValues = new object[rows, cols]; // 0-based array
                    for (int r = 0; r < rows; r++)
                    {
                        for (int c = 0; c < cols; c++)
                        {
                            // Convert JsonElement to proper C# type for COM interop
                            // MCP framework deserializes JSON to JsonElement, not primitives
                            arrayValues[r, c] = RangeHelpers.ConvertToCellValue(values[r][c]);
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
                result.Success = false;
                result.ErrorMessage = $"Cannot write to range '{rangeAddress}' on sheet '{sheetName}': {comEx.Message}";
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref range);
            }
        });
    }
}
