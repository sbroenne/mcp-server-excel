using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet lifecycle operations (List, Create, Rename, Copy, Delete)
/// </summary>
public partial class SheetCommands
{
    /// <inheritdoc />
    public WorksheetListResult List(IExcelBatch batch)
    {
        var result = new WorksheetListResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
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
    /// <inheritdoc />
    public OperationResult Create(IExcelBatch batch, string sheetName)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "create-sheet" };

        return batch.Execute((ctx, ct) =>
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
    /// <inheritdoc />
    public OperationResult Rename(IExcelBatch batch, string oldName, string newName)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "rename-sheet" };

        return batch.Execute((ctx, ct) =>
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
    /// <inheritdoc />
    public OperationResult Copy(IExcelBatch batch, string sourceName, string targetName)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "copy-sheet" };

        return batch.Execute((ctx, ct) =>
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
    /// <inheritdoc />
    public OperationResult Delete(IExcelBatch batch, string sheetName)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "delete-sheet" };

        return batch.Execute((ctx, ct) =>
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
    public OperationResult Move(IExcelBatch batch, string sheetName, string? beforeSheet = null, string? afterSheet = null)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "move-sheet" };

        // Validate parameters
        if (!string.IsNullOrWhiteSpace(beforeSheet) && !string.IsNullOrWhiteSpace(afterSheet))
        {
            result.Success = false;
            result.ErrorMessage = "Cannot specify both beforeSheet and afterSheet";
            return result;
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? targetSheet = null;
            dynamic? sheets = null;
            dynamic? lastSheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return result;
                }

                // If no position specified, move to end
                if (string.IsNullOrWhiteSpace(beforeSheet) && string.IsNullOrWhiteSpace(afterSheet))
                {
                    sheets = ctx.Book.Worksheets;
                    lastSheet = sheets.Item(sheets.Count);
                    sheet.Move(After: lastSheet);
                }
                else
                {
                    // Find target sheet for positioning
                    string targetName = beforeSheet ?? afterSheet!;
                    targetSheet = ComUtilities.FindSheet(ctx.Book, targetName);
                    if (targetSheet == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Target sheet '{targetName}' not found";
                        return result;
                    }

                    // Move using Excel COM API
                    if (!string.IsNullOrWhiteSpace(beforeSheet))
                    {
                        sheet.Move(Before: targetSheet);
                    }
                    else
                    {
                        sheet.Move(After: targetSheet);
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
                ComUtilities.Release(ref lastSheet);
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref targetSheet);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult CopyToWorkbook(IExcelBatch sourceBatch, string sourceSheet, IExcelBatch targetBatch, string? targetSheetName = null, string? beforeSheet = null, string? afterSheet = null)
    {
        // Cross-workbook operations not currently supported due to COM RCW limitations
        // The IExcelBatch pattern isolates each batch.Execute() context, preventing COM objects
        // from crossing batch boundaries. Supporting this requires architectural changes.
        var result = new OperationResult
        {
            FilePath = sourceBatch.WorkbookPath,
            Action = "copy-to-workbook",
            Success = false,
            ErrorMessage = "Cross-workbook sheet operations are not currently supported due to COM interop architecture limitations. " +
                           "Workaround: Open both workbooks in Excel and use Copy Sheet manually, or export/import the sheet data."
        };
        return result;
    }

    /// <inheritdoc />
    public OperationResult MoveToWorkbook(IExcelBatch sourceBatch, string sourceSheet, IExcelBatch targetBatch, string? beforeSheet = null, string? afterSheet = null)
    {
        // Cross-workbook operations not currently supported due to COM RCW limitations
        // The IExcelBatch pattern isolates each batch.Execute() context, preventing COM objects
        // from crossing batch boundaries. Supporting this requires architectural changes.
        var result = new OperationResult
        {
            FilePath = sourceBatch.WorkbookPath,
            Action = "move-to-workbook",
            Success = false,
            ErrorMessage = "Cross-workbook sheet operations are not currently supported due to COM interop architecture limitations. " +
                           "Workaround: Open both workbooks in Excel and use Move/Copy Sheet manually, or copy data between workbooks."
        };
        return result;
    }
}
