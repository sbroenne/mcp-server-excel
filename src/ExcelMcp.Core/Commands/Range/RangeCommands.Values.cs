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
                originalCalculation = (int)ctx.App.Calculation;
                if (originalCalculation != -4135) // xlCalculationManual
                {
                    ctx.App.Calculation = (Excel.XlCalculation)(-4135); // xlCalculationManual
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
                // Restore original calculation mode
                if (calculationChanged && originalCalculation != -1)
                {
                    try
                    {
                        ctx.App.Calculation = (Excel.XlCalculation)originalCalculation;
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        // Ignore errors restoring calculation mode - not critical
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
}



