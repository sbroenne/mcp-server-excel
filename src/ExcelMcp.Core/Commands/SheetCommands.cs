using Sbroenne.ExcelMcp.Core.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Session;

#pragma warning disable CS1998 // Async method lacks 'await' operators - intentional for COM synchronous operations

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet management commands implementation - all operations use batching for performance.
/// </summary>
public class SheetCommands : ISheetCommands
{
    /// <inheritdoc />
    public async Task<WorksheetListResult> ListAsync(IExcelBatch batch)
    {
        var result = new WorksheetListResult { FilePath = batch.WorkbookPath };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? sheets = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic? sheet = null;
                    try
                    {
                        sheet = sheets.Item(i);
                        result.Worksheets.Add(new WorksheetInfo { Name = sheet.Name, Index = i });
                    }
                    finally
                    {
                        ComUtilities.Release(ref sheet);
                    }
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
                ComUtilities.Release(ref sheets);
            }
        });
    }

    /// <inheritdoc />
    public async Task<WorksheetDataResult> ReadAsync(IExcelBatch batch, string sheetName, string? range = null)
    {
        var result = new WorksheetDataResult { FilePath = batch.WorkbookPath, SheetName = sheetName, Range = range ?? string.Empty };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? rangeObj = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return result;
                }

                if (string.IsNullOrWhiteSpace(range))
                {
                    rangeObj = sheet.UsedRange;
                }
                else
                {
                    rangeObj = sheet.Range[range];
                }

                object[,] values = rangeObj.Value2;
                if (values != null)
                {
                    int rows = values.GetLength(0), cols = values.GetLength(1);
                    for (int r = 1; r <= rows; r++)
                    {
                        var row = new List<object?>();
                        for (int c = 1; c <= cols; c++) row.Add(values[r, c]);
                        result.Data.Add(row);
                    }
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
                ComUtilities.Release(ref rangeObj);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> WriteAsync(IExcelBatch batch, string sheetName, string csvData)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "write" };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? cell1 = null;
            dynamic? cell2 = null;
            dynamic? range = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return result;
                }

                var data = ParseCsv(csvData);
                if (data.Count == 0)
                {
                    result.Success = false;
                    result.ErrorMessage = "No data to write";
                    return result;
                }

                int rows = data.Count, cols = data[0].Count;
                object[,] arr = new object[rows, cols];
                for (int r = 0; r < rows; r++)
                    for (int c = 0; c < cols; c++)
                        arr[r, c] = data[r][c];

                cell1 = sheet.Cells[1, 1];
                cell2 = sheet.Cells[rows, cols];
                range = sheet.Range[cell1, cell2];
                range.Value2 = arr;

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
                ComUtilities.Release(ref cell2);
                ComUtilities.Release(ref cell1);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> CreateAsync(IExcelBatch batch, string sheetName)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "create-sheet" };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? sheets = null;
            dynamic? newSheet = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                newSheet = sheets.Add();
                newSheet.Name = sheetName;
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
                ComUtilities.Release(ref newSheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> RenameAsync(IExcelBatch batch, string oldName, string newName)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "rename-sheet" };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, oldName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{oldName}' not found";
                    return result;
                }
                sheet.Name = newName;
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
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> CopyAsync(IExcelBatch batch, string sourceName, string targetName)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "copy-sheet" };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? sourceSheet = null;
            dynamic? sheets = null;
            dynamic? lastSheet = null;
            dynamic? copiedSheet = null;
            try
            {
                sourceSheet = ComUtilities.FindSheet(ctx.Book, sourceName);
                if (sourceSheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sourceName}' not found";
                    return result;
                }
                sheets = ctx.Book.Worksheets;
                lastSheet = sheets.Item(sheets.Count);
                sourceSheet.Copy(After: lastSheet);
                copiedSheet = sheets.Item(sheets.Count);
                copiedSheet.Name = targetName;
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
                ComUtilities.Release(ref copiedSheet);
                ComUtilities.Release(ref lastSheet);
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref sourceSheet);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> DeleteAsync(IExcelBatch batch, string sheetName)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "delete-sheet" };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return result;
                }
                sheet.Delete();
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
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> ClearAsync(IExcelBatch batch, string sheetName, string? range = null)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "clear" };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? rangeObj = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return result;
                }
                rangeObj = range != null ? sheet.Range[range] : sheet.UsedRange;
                rangeObj.Clear();
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
                ComUtilities.Release(ref rangeObj);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> AppendAsync(IExcelBatch batch, string sheetName, string csvData)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "append" };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? usedRange = null;
            dynamic? rows = null;
            dynamic? cell1 = null;
            dynamic? cell2 = null;
            dynamic? range = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return result;
                }

                usedRange = sheet.UsedRange;
                rows = usedRange.Rows;
                int lastRow = rows.Count;

                var data = ParseCsv(csvData);
                if (data.Count == 0)
                {
                    result.Success = false;
                    result.ErrorMessage = "No data to append";
                    return result;
                }

                int startRow = lastRow + 1, numRows = data.Count, cols = data[0].Count;
                object[,] arr = new object[numRows, cols];
                for (int r = 0; r < numRows; r++)
                    for (int c = 0; c < cols; c++)
                        arr[r, c] = data[r][c];

                cell1 = sheet.Cells[startRow, 1];
                cell2 = sheet.Cells[startRow + numRows - 1, cols];
                range = sheet.Range[cell1, cell2];
                range.Value2 = arr;
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
                ComUtilities.Release(ref cell2);
                ComUtilities.Release(ref cell1);
                ComUtilities.Release(ref rows);
                ComUtilities.Release(ref usedRange);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static List<List<string>> ParseCsv(string csvData)
    {
        var result = new List<List<string>>();
        var lines = csvData.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var line in lines)
        {
            var row = new List<string>();
            var fields = line.Split(',');
            foreach (var field in fields)
                row.Add(field.Trim().Trim('"'));
            result.Add(row);
        }
        return result;
    }
}
