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
    public WorksheetListResult List(IExcelBatch batch, string? filePath = null)
    {
        var result = new WorksheetListResult { FilePath = filePath ?? batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            // Get the workbook to list from
            dynamic workbook = filePath != null ? batch.GetWorkbook(filePath) : ctx.Book;

            dynamic? sheets = null;
            try
            {
                sheets = workbook.Worksheets;
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
            finally
            {
                ComUtilities.Release(ref sheets);
            }
        });
    }

    /// <inheritdoc />
    /// <inheritdoc />
    public OperationResult Create(IExcelBatch batch, string sheetName, string? filePath = null)
    {
        var result = new OperationResult { FilePath = filePath ?? batch.WorkbookPath, Action = "create-sheet" };

        return batch.Execute((ctx, ct) =>
        {
            // Get the workbook to create sheet in
            dynamic workbook = filePath != null ? batch.GetWorkbook(filePath) : ctx.Book;

            dynamic? sheets = null;
            dynamic? newSheet = null;
            try
            {
                sheets = workbook.Worksheets;
                newSheet = sheets.Add();
                newSheet.Name = sheetName;
                result.Success = true;
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
    public OperationResult CopyToWorkbook(IExcelBatch batch, string sourceFile, string sourceSheet, string targetFile, string? targetSheetName = null, string? beforeSheet = null, string? afterSheet = null)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "copy-to-workbook"
        };

        // Validate positioning parameters
        if (!string.IsNullOrWhiteSpace(beforeSheet) && !string.IsNullOrWhiteSpace(afterSheet))
        {
            result.Success = false;
            result.ErrorMessage = "Cannot specify both beforeSheet and afterSheet. Choose one or neither.";
            return result;
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic? sourceWb = null;
            dynamic? targetWb = null;
            dynamic? sourceSheetObj = null;
            dynamic? targetSheets = null;
            dynamic? targetPositionSheet = null;
            dynamic? copiedSheet = null;

            try
            {
                // Get both workbooks from the batch
                sourceWb = batch.GetWorkbook(sourceFile);
                targetWb = batch.GetWorkbook(targetFile);

                // Find source sheet
                sourceSheetObj = ComUtilities.FindSheet(sourceWb, sourceSheet);
                if (sourceSheetObj == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Source sheet '{sourceSheet}' not found in '{Path.GetFileName(sourceFile)}'";
                    return result;
                }

                // Handle positioning
                targetSheets = targetWb.Worksheets;

                if (!string.IsNullOrWhiteSpace(beforeSheet))
                {
                    targetPositionSheet = ComUtilities.FindSheet(targetWb, beforeSheet);
                    if (targetPositionSheet == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Target sheet '{beforeSheet}' not found in '{Path.GetFileName(targetFile)}'";
                        return result;
                    }
                    // Copy before specified sheet
                    sourceSheetObj.Copy(Before: targetPositionSheet);
                }
                else if (!string.IsNullOrWhiteSpace(afterSheet))
                {
                    targetPositionSheet = ComUtilities.FindSheet(targetWb, afterSheet);
                    if (targetPositionSheet == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Target sheet '{afterSheet}' not found in '{Path.GetFileName(targetFile)}'";
                        return result;
                    }
                    // Copy after specified sheet
                    sourceSheetObj.Copy(After: targetPositionSheet);
                }
                else
                {
                    // Copy to end of target workbook
                    dynamic? lastSheet = targetSheets.Item(targetSheets.Count);
                    try
                    {
                        sourceSheetObj.Copy(After: lastSheet);
                    }
                    finally
                    {
                        ComUtilities.Release(ref lastSheet!);
                    }
                }

                // Rename if requested
                if (!string.IsNullOrWhiteSpace(targetSheetName))
                {
                    copiedSheet = targetSheets.Item(targetSheets.Count); // Last sheet is the copied one
                    copiedSheet.Name = targetSheetName;
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref copiedSheet);
                ComUtilities.Release(ref targetPositionSheet);
                ComUtilities.Release(ref targetSheets);
                ComUtilities.Release(ref sourceSheetObj);
                // Note: Don't release sourceWb and targetWb - they're managed by the batch
            }
        });
    }

    /// <inheritdoc />
    public OperationResult MoveToWorkbook(IExcelBatch batch, string sourceFile, string sourceSheet, string targetFile, string? beforeSheet = null, string? afterSheet = null)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "move-to-workbook"
        };

        // Validate positioning parameters
        if (!string.IsNullOrWhiteSpace(beforeSheet) && !string.IsNullOrWhiteSpace(afterSheet))
        {
            result.Success = false;
            result.ErrorMessage = "Cannot specify both beforeSheet and afterSheet. Choose one or neither.";
            return result;
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic? sourceWb = null;
            dynamic? targetWb = null;
            dynamic? sourceSheetObj = null;
            dynamic? targetSheets = null;
            dynamic? targetPositionSheet = null;

            try
            {
                // Get both workbooks from the batch
                sourceWb = batch.GetWorkbook(sourceFile);
                targetWb = batch.GetWorkbook(targetFile);

                // Find source sheet
                sourceSheetObj = ComUtilities.FindSheet(sourceWb, sourceSheet);
                if (sourceSheetObj == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Source sheet '{sourceSheet}' not found in '{Path.GetFileName(sourceFile)}'";
                    return result;
                }

                // Handle positioning
                targetSheets = targetWb.Worksheets;

                if (!string.IsNullOrWhiteSpace(beforeSheet))
                {
                    targetPositionSheet = ComUtilities.FindSheet(targetWb, beforeSheet);
                    if (targetPositionSheet == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Target sheet '{beforeSheet}' not found in '{Path.GetFileName(targetFile)}'";
                        return result;
                    }
                    // Move before specified sheet
                    sourceSheetObj.Move(Before: targetPositionSheet);
                }
                else if (!string.IsNullOrWhiteSpace(afterSheet))
                {
                    targetPositionSheet = ComUtilities.FindSheet(targetWb, afterSheet);
                    if (targetPositionSheet == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Target sheet '{afterSheet}' not found in '{Path.GetFileName(targetFile)}'";
                        return result;
                    }
                    // Move after specified sheet
                    sourceSheetObj.Move(After: targetPositionSheet);
                }
                else
                {
                    // Move to end of target workbook
                    dynamic? lastSheet = targetSheets.Item(targetSheets.Count);
                    try
                    {
                        sourceSheetObj.Move(After: lastSheet);
                    }
                    finally
                    {
                        ComUtilities.Release(ref lastSheet!);
                    }
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref targetPositionSheet);
                ComUtilities.Release(ref targetSheets);
                ComUtilities.Release(ref sourceSheetObj);
                // Note: Don't release sourceWb and targetWb - they're managed by the batch
            }
        });
    }
}
