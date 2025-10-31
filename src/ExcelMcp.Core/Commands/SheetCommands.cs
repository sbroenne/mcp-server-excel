using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;


namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet lifecycle management commands - create, rename, copy, delete worksheets.
/// Data operations (read, write, clear) moved to Range.IRangeCommands for unified range API.
/// All operations use batching for performance.
/// </summary>
public class SheetCommands : ISheetCommands
{
    /// <inheritdoc />
    public async Task<WorksheetListResult> ListAsync(IExcelBatch batch)
    {
        var result = new WorksheetListResult { FilePath = batch.WorkbookPath };

        return await batch.Execute((ctx, ct) =>
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
    public async Task<OperationResult> CreateAsync(IExcelBatch batch, string sheetName)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "create-sheet" };

        return await batch.Execute((ctx, ct) =>
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

        return await batch.Execute((ctx, ct) =>
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

        return await batch.Execute((ctx, ct) =>
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

        return await batch.Execute((ctx, ct) =>
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
}
