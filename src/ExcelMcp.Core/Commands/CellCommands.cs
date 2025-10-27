using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop.Session;

#pragma warning disable CS1998 // Async method lacks 'await' operators - intentional for COM synchronous operations

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Individual cell operation commands implementation
/// </summary>
public class CellCommands : ICellCommands
{
    /// <inheritdoc />
    public async Task<CellValueResult> GetValueAsync(IExcelBatch batch, string sheetName, string cellAddress)
    {
        var result = new CellValueResult
        {
            FilePath = batch.WorkbookPath,
            CellAddress = $"{sheetName}!{cellAddress}"
        };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? cell = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return result;
                }

                cell = sheet.Range[cellAddress];
                result.Value = cell.Value2;
                result.ValueType = result.Value?.GetType().Name ?? "null";
                result.Formula = cell.Formula;
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
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> SetValueAsync(IExcelBatch batch, string sheetName, string cellAddress, string value)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "set-value"
        };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? cell = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return result;
                }

                cell = sheet.Range[cellAddress];

                // Try to parse as number, otherwise set as text
                if (double.TryParse(value, out double numValue))
                {
                    cell.Value2 = numValue;
                }
                else if (bool.TryParse(value, out bool boolValue))
                {
                    cell.Value2 = boolValue;
                }
                else
                {
                    cell.Value2 = value;
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
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public async Task<CellValueResult> GetFormulaAsync(IExcelBatch batch, string sheetName, string cellAddress)
    {
        var result = new CellValueResult
        {
            FilePath = batch.WorkbookPath,
            CellAddress = $"{sheetName}!{cellAddress}"
        };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? cell = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return result;
                }

                cell = sheet.Range[cellAddress];
                result.Formula = cell.Formula ?? "";
                result.Value = cell.Value2;
                result.ValueType = result.Value?.GetType().Name ?? "null";
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
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> SetFormulaAsync(IExcelBatch batch, string sheetName, string cellAddress, string formula)
    {
        // Ensure formula starts with =
        if (!formula.StartsWith("="))
        {
            formula = "=" + formula;
        }

        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "set-formula"
        };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? cell = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return result;
                }

                cell = sheet.Range[cellAddress];
                cell.Formula = formula;

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
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
    }
}
