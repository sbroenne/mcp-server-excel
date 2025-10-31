using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;


namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Range formula operations (get/set formulas as 2D arrays)
/// </summary>
public partial class RangeCommands
{
    /// <inheritdoc />
    public async Task<RangeFormulaResult> GetFormulasAsync(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        var result = new RangeFormulaResult
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
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress);
                if (range == null)
                {
                    result.Success = false;
                    result.ErrorMessage = RangeHelpers.GetResolveError(sheetName, rangeAddress);
                    return result;
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
                            formulaRow.Add(formula.StartsWith("=") ? formula : string.Empty);
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
                    result.Formulas.Add([formula.StartsWith("=") ? formula : string.Empty]);
                    result.Values.Add([valueOrArray]);
                }

                result.Success = true;
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
    public async Task<OperationResult> SetFormulasAsync(IExcelBatch batch, string sheetName, string rangeAddress, List<List<string>> formulas)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "set-formulas" };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress);
                if (range == null)
                {
                    result.Success = false;
                    result.ErrorMessage = RangeHelpers.GetResolveError(sheetName, rangeAddress);
                    return result;
                }

                // Convert List<List<string>> to 2D array (0-based for Excel COM)
                int rows = formulas.Count;
                int cols = formulas.Count > 0 ? formulas[0].Count : 0;

                if (rows > 0 && cols > 0)
                {
                    object[,] arrayFormulas = new object[rows, cols]; // 0-based array
                    for (int r = 0; r < rows; r++)
                    {
                        for (int c = 0; c < cols; c++)
                        {
                            // Convert JsonElement to proper C# type for COM interop
                            // MCP framework deserializes JSON to JsonElement, not primitives
                            arrayFormulas[r, c] = RangeHelpers.ConvertToCellValue(formulas[r][c]);
                        }
                    }

                    range.Formula = arrayFormulas;
                }

                result.Success = true;
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
