using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;


namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Range formula operations (get/set formulas as 2D arrays)
/// </summary>
public partial class RangeCommands
{
    /// <inheritdoc />
    public RangeFormulaResult GetFormulas(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        var result = new RangeFormulaResult
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

                // Get actual address
                result.RangeAddress = range.Address;

                // Get formulas and values - handle single cell case
                object formulaOrArray = range.Formula;
                object valueOrArray = range.Value2;

                if (formulaOrArray is object[,] formulas && valueOrArray is object[,] values)
                {
                    // Multi-cell range
                    result.RowCount = formulas.GetLength(0);
                    result.ColumnCount = formulas.GetLength(1);

                    for (int r = 1; r <= result.RowCount; r++)
                    {
                        var formulaRow = new List<string>();
                        var valueRow = new List<object?>();

                        for (int c = 1; c <= result.ColumnCount; c++)
                        {
                            string formula = formulas[r, c]?.ToString() ?? string.Empty;
                            // Only return actual formulas (starting with =), not values
                            formulaRow.Add(formula.StartsWith('=') ? formula : string.Empty);
                            valueRow.Add(values[r, c]);
                        }

                        result.Formulas.Add(formulaRow);
                        result.Values.Add(valueRow);
                    }
                }
                else
                {
                    // Single cell
                    result.RowCount = 1;
                    result.ColumnCount = 1;
                    string formula = formulaOrArray?.ToString() ?? string.Empty;
                    // Only return actual formulas (starting with =), not values
                    result.Formulas.Add([formula.StartsWith('=') ? formula : string.Empty]);
                    result.Values.Add([valueOrArray]);
                }

                result.Success = true;
                return result;
            }
            catch (System.Runtime.InteropServices.COMException comEx) when (comEx.HResult == unchecked((int)0x8007000E))
            {
                // E_OUTOFMEMORY - Excel's misleading error for sheet/range/session issues
                throw new InvalidOperationException($"Cannot read formulas from range '{rangeAddress}' on sheet '{sheetName}': {comEx.Message}", comEx);
            }
            finally
            {
                ComUtilities.Release(ref range);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetFormulas(IExcelBatch batch, string sheetName, string rangeAddress, List<List<string>>? formulas = null, string? formulasFile = null)
    {
        // Resolve formulas from inline parameter or file
        var resolvedFormulas = ParameterTransforms.ResolveFormulasOrFile(formulas, formulasFile);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "set-formulas" };

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
                // hanging when formulas reference Data Model/DAX query tables or complex calculations.
                // Without this, setting formulas that trigger recalculation can block the COM interface.
                originalCalculation = ctx.App.Calculation;
                if (originalCalculation != -4135) // xlCalculationManual
                {
                    ctx.App.Calculation = -4135; // xlCalculationManual
                    calculationChanged = true;
                }

                // Convert List<List<string>> to 2D array
                // Excel COM requires 1-based arrays for multi-cell ranges
                int rows = resolvedFormulas.Count;
                int cols = resolvedFormulas.Count > 0 ? resolvedFormulas[0].Count : 0;

                if (rows > 0 && cols > 0)
                {
                    // Create 1-based array for Excel COM compatibility
                    object[,] arrayFormulas = (object[,])Array.CreateInstance(typeof(object), [rows, cols], [1, 1]);

                    for (int r = 1; r <= rows; r++)
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            // Convert JsonElement to proper C# type for COM interop
                            // MCP framework deserializes JSON to JsonElement, not primitives
                            arrayFormulas[r, c] = RangeHelpers.ConvertToCellValue(resolvedFormulas[r - 1][c - 1]);
                        }
                    }

                    range.Formula = arrayFormulas;
                }

                result.Success = true;
                return result;
            }
            catch (System.Runtime.InteropServices.COMException comEx) when (comEx.HResult == unchecked((int)0x8007000E))
            {
                // E_OUTOFMEMORY - Excel's misleading error for sheet/range/session issues
                throw new InvalidOperationException($"Cannot write formulas to range '{rangeAddress}' on sheet '{sheetName}': {comEx.Message}", comEx);
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
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        // Ignore errors restoring calculation mode - not critical
                    }
                }
                ComUtilities.Release(ref range);
            }
        });
    }
}



