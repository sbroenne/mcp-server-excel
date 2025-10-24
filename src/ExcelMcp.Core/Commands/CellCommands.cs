using Sbroenne.ExcelMcp.Core.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Session;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Individual cell operation commands implementation
/// </summary>
public class CellCommands : ICellCommands
{
    /// <inheritdoc />
    public CellValueResult GetValue(string filePath, string sheetName, string cellAddress)
    {
        if (!File.Exists(filePath))
        {
            return new CellValueResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath,
                CellAddress = cellAddress
            };
        }

        var result = new CellValueResult
        {
            FilePath = filePath,
            CellAddress = $"{sheetName}!{cellAddress}"
        };

        ExcelSession.Execute(filePath, false, (excel, workbook) =>
        {
            dynamic? sheet = null;
            dynamic? cell = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return 1;
                }

                cell = sheet.Range[cellAddress];
                result.Value = cell.Value2;
                result.ValueType = result.Value?.GetType().Name ?? "null";
                result.Formula = cell.Formula;
                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return 1;
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult SetValue(string filePath, string sheetName, string cellAddress, string value)
    {
        if (!File.Exists(filePath))
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath,
                Action = "set-value"
            };
        }

        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "set-value"
        };

        ExcelSession.Execute(filePath, true, (excel, workbook) =>
        {
            dynamic? sheet = null;
            dynamic? cell = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return 1;
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

                workbook.Save();
                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return 1;
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });

        return result;
    }

    /// <inheritdoc />
    public CellValueResult GetFormula(string filePath, string sheetName, string cellAddress)
    {
        if (!File.Exists(filePath))
        {
            return new CellValueResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath,
                CellAddress = cellAddress
            };
        }

        var result = new CellValueResult
        {
            FilePath = filePath,
            CellAddress = $"{sheetName}!{cellAddress}"
        };

        ExcelSession.Execute(filePath, false, (excel, workbook) =>
        {
            dynamic? sheet = null;
            dynamic? cell = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return 1;
                }

                cell = sheet.Range[cellAddress];
                result.Formula = cell.Formula ?? "";
                result.Value = cell.Value2;
                result.ValueType = result.Value?.GetType().Name ?? "null";
                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return 1;
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult SetFormula(string filePath, string sheetName, string cellAddress, string formula)
    {
        if (!File.Exists(filePath))
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath,
                Action = "set-formula"
            };
        }

        // Ensure formula starts with =
        if (!formula.StartsWith("="))
        {
            formula = "=" + formula;
        }

        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "set-formula"
        };

        ExcelSession.Execute(filePath, true, (excel, workbook) =>
        {
            dynamic? sheet = null;
            dynamic? cell = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return 1;
                }

                cell = sheet.Range[cellAddress];
                cell.Formula = formula;

                workbook.Save();
                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return 1;
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });

        return result;
    }
}
