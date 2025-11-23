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
    public void Create(IExcelBatch batch, string sheetName, string? filePath = null)
    {
        batch.Execute((ctx, ct) =>
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
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref newSheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    /// <inheritdoc />
    public void Rename(IExcelBatch batch, string oldName, string newName)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, oldName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{oldName}' not found");
                }
                sheet.Name = newName;
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public void Copy(IExcelBatch batch, string sourceName, string targetName)
    {
        batch.Execute((ctx, ct) =>
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
                    throw new InvalidOperationException($"Sheet '{sourceName}' not found");
                }
                sheets = ctx.Book.Worksheets;
                lastSheet = sheets.Item(sheets.Count);
                sourceSheet.Copy(After: lastSheet);
                copiedSheet = sheets.Item(sheets.Count);
                copiedSheet.Name = targetName;
                return 0;
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
    public void Delete(IExcelBatch batch, string sheetName)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found");
                }
                sheet.Delete();
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public void Move(IExcelBatch batch, string sheetName, string? beforeSheet = null, string? afterSheet = null)
    {
        // Validate parameters
        if (!string.IsNullOrWhiteSpace(beforeSheet) && !string.IsNullOrWhiteSpace(afterSheet))
        {
            throw new ArgumentException("Cannot specify both beforeSheet and afterSheet");
        }

        batch.Execute((ctx, ct) =>
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
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found");
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
                        throw new InvalidOperationException($"Target sheet '{targetName}' not found");
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
                return 0;
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
    public void CopyToWorkbook(IExcelBatch batch, string sourceFile, string sourceSheet, string targetFile, string? targetSheetName = null, string? beforeSheet = null, string? afterSheet = null)
    {
        // Validate positioning parameters
        if (!string.IsNullOrWhiteSpace(beforeSheet) && !string.IsNullOrWhiteSpace(afterSheet))
        {
            throw new ArgumentException("Cannot specify both beforeSheet and afterSheet. Choose one or neither.");
        }

        batch.Execute((ctx, ct) =>
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
                    throw new InvalidOperationException($"Source sheet '{sourceSheet}' not found in '{Path.GetFileName(sourceFile)}'");
                }

                // Handle positioning
                targetSheets = targetWb.Worksheets;

                if (!string.IsNullOrWhiteSpace(beforeSheet))
                {
                    targetPositionSheet = ComUtilities.FindSheet(targetWb, beforeSheet);
                    if (targetPositionSheet == null)
                    {
                        throw new InvalidOperationException($"Target sheet '{beforeSheet}' not found in '{Path.GetFileName(targetFile)}'");
                    }
                    // Copy before specified sheet
                    sourceSheetObj.Copy(Before: targetPositionSheet);
                }
                else if (!string.IsNullOrWhiteSpace(afterSheet))
                {
                    targetPositionSheet = ComUtilities.FindSheet(targetWb, afterSheet);
                    if (targetPositionSheet == null)
                    {
                        throw new InvalidOperationException($"Target sheet '{afterSheet}' not found in '{Path.GetFileName(targetFile)}'");
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
                return 0;
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
    public void MoveToWorkbook(IExcelBatch batch, string sourceFile, string sourceSheet, string targetFile, string? beforeSheet = null, string? afterSheet = null)
    {
        // Validate positioning parameters
        if (!string.IsNullOrWhiteSpace(beforeSheet) && !string.IsNullOrWhiteSpace(afterSheet))
        {
            throw new ArgumentException("Cannot specify both beforeSheet and afterSheet. Choose one or neither.");
        }

        batch.Execute((ctx, ct) =>
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
                    throw new InvalidOperationException($"Source sheet '{sourceSheet}' not found in '{Path.GetFileName(sourceFile)}'");
                }

                // Handle positioning
                targetSheets = targetWb.Worksheets;

                if (!string.IsNullOrWhiteSpace(beforeSheet))
                {
                    targetPositionSheet = ComUtilities.FindSheet(targetWb, beforeSheet);
                    if (targetPositionSheet == null)
                    {
                        throw new InvalidOperationException($"Target sheet '{beforeSheet}' not found in '{Path.GetFileName(targetFile)}'");
                    }
                    // Move before specified sheet
                    sourceSheetObj.Move(Before: targetPositionSheet);
                }
                else if (!string.IsNullOrWhiteSpace(afterSheet))
                {
                    targetPositionSheet = ComUtilities.FindSheet(targetWb, afterSheet);
                    if (targetPositionSheet == null)
                    {
                        throw new InvalidOperationException($"Target sheet '{afterSheet}' not found in '{Path.GetFileName(targetFile)}'");
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
                return 0;
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
