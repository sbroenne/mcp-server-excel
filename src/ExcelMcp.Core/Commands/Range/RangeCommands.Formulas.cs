using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;


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
            RangeAddress = rangeAddress,
            CellErrors = []
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
                // Use Formula2 (modern) instead of Formula (legacy) to avoid implicit intersection (@)
                // operator being injected in Excel Table cells. Formula2 respects dynamic array semantics.
                object formulaOrArray = range.Formula2;
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
                            object? cellValue = values[r, c];

                            // Only return actual formulas (starting with =), not values
                            formulaRow.Add(formula.StartsWith('=') ? formula : string.Empty);
                            valueRow.Add(cellValue);

                            // ERROR CODE DETECTION: Map Excel error codes to human-readable messages
                            if (cellValue is int errorCode && errorCode < 0)
                            {
                                var cellAddr = $"{GetColumnLetter(c)}{r}";
                                result.CellErrors.Add(new RangeCellError
                                {
                                    CellAddress = cellAddr,
                                    ErrorCode = errorCode,
                                    ErrorMessage = MapErrorCodeToMessage(errorCode)
                                });
                            }
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
                    object? cellValue = valueOrArray;

                    // Only return actual formulas (starting with =), not values
                    result.Formulas.Add([formula.StartsWith('=') ? formula : string.Empty]);
                    result.Values.Add([cellValue]);

                    // ERROR CODE DETECTION: Single cell error
                    if (cellValue is int errorCode && errorCode < 0)
                    {
                        result.CellErrors.Add(new RangeCellError
                        {
                            CellAddress = range.Address,
                            ErrorCode = errorCode,
                            ErrorMessage = MapErrorCodeToMessage(errorCode)
                        });
                    }
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

    /// <summary>
    /// Maps Excel error codes to human-readable error messages
    /// </summary>
    private static string MapErrorCodeToMessage(int errorCode) =>
        errorCode switch
        {
            -2146826288 => "#NULL! - Invalid intersection of ranges",
            -2147483648 => "#DIV/0! - Division by zero",
            -2146826259 => "#VALUE! - Wrong type of argument",
            -2146826246 => "#REF! - Invalid cell reference",
            -2146826252 => "#NUM! - Invalid numeric value",
            -2142019887 => "#N/A - Value not available",
            _ => $"#ERROR! - Unknown error code {errorCode}"
        };

    /// <summary>
    /// Converts 1-based column index to Excel column letter (1=A, 26=Z, 27=AA)
    /// </summary>
    private static string GetColumnLetter(int columnIndex)
    {
        string columnName = string.Empty;
        while (columnIndex > 0)
        {
            columnIndex--;
            columnName = Convert.ToChar('A' + (columnIndex % 26)) + columnName;
            columnIndex /= 26;
        }
        return columnName;
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

                // Convert List<List<string>> to 2D array
                // Excel COM requires 1-based arrays for multi-cell ranges
                int rows = resolvedFormulas.Count;
                int cols = resolvedFormulas.Count > 0 ? resolvedFormulas[0].Count : 0;

                ValidateRectangularRowWidths(resolvedFormulas, Convert.ToInt32(range.Columns.Count), nameof(formulas), "Formula");

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

                    // Use Formula2 (modern) instead of Formula (legacy) to prevent Excel from
                    // injecting the @ implicit intersection operator in table cells, which causes
                    // #FIELD! errors with custom functions that return entity cards.
                    range.Formula2 = arrayFormulas;
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
}



