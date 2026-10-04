using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

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
                        result.Worksheets.Add(new WorksheetInfo
                        {
                            Name = sheet.Name,
                            Index = i,
                            Visible = (int)sheet.Visible == -1  // xlSheetVisible = -1
                        });
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
    public OperationResult Create(IExcelBatch batch, string sheetName, string? filePath = null)
    {
        ValidateNewSheetName(sheetName);
        return batch.Execute((ctx, ct) =>
        {
            // Get the workbook to create sheet in
            Excel.Workbook workbook = filePath != null ? batch.GetWorkbook(filePath) : ctx.Book;

            Excel.Sheets? sheets = null;
            Excel.Worksheet? newSheet = null;
            try
            {
                EnsureSheetNameAvailable(workbook, sheetName, ct);
                sheets = workbook.Worksheets;
                newSheet = (Excel.Worksheet)sheets.Add();
                var addedSheetCurrentName = newSheet.Name;
                SetSheetNameWithContext(newSheet, sheetName, addedSheetCurrentName, isNewSheet: true);
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref newSheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Rename(IExcelBatch batch, string oldName, string newName)
    {
        ValidateNewSheetName(newName);
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, oldName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{oldName}' not found.");
                }
                EnsureSheetNameAvailable(ctx.Book, newName, ct, sheet.Name);
                SetSheetNameWithContext(sheet, newName, sheet.Name, isNewSheet: false);
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Copy(IExcelBatch batch, string sourceName, string targetName)
    {
        ValidateNewSheetName(targetName);
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sourceSheet = null;
            Excel.Sheets? sheets = null;
            Excel.Worksheet? lastSheet = null;
            Excel.Worksheet? copiedSheet = null;
            try
            {
                sourceSheet = ComUtilities.FindSheet(ctx.Book, sourceName);
                if (sourceSheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sourceName}' not found.");
                }
                EnsureSheetNameAvailable(ctx.Book, targetName, ct);
                sheets = ctx.Book.Worksheets;
                lastSheet = (Excel.Worksheet)sheets.Item[sheets.Count];
                sourceSheet.Copy(After: lastSheet);
                copiedSheet = (Excel.Worksheet)sheets.Item[sheets.Count];
                var copiedSheetCurrentName = copiedSheet.Name;
                SetSheetNameWithContext(copiedSheet, targetName, copiedSheetCurrentName, isNewSheet: true);
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
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

    private static void SetSheetNameWithContext(
        Excel.Worksheet sheet, string name, string existingName, bool isNewSheet)
    {
        try
        {
            sheet.Name = name;
        }
        catch (System.Runtime.InteropServices.COMException namingError)
        {
            var message = isNewSheet
                ? $"Excel rejected worksheet name '{name}'. The new worksheet '{existingName}' remains in the workbook; inspect it and remove it if appropriate."
                : $"Excel rejected worksheet name '{name}'. Worksheet '{existingName}' was not renamed.";
            throw new InvalidOperationException(message, namingError);
        }
    }

    private static void ValidateNewSheetName(string name)
    {
        if (string.IsNullOrWhiteSpace(name) ||
            name.Length > 31 ||
            name.AsSpan().IndexOfAny(":\\/?*[]".AsSpan()) >= 0 ||
            name.StartsWith('\'') ||
            name.EndsWith('\'') ||
            string.Equals(name, "History", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException(
                "Worksheet names must be nonblank, contain at most 31 characters, not be 'History', " +
                "not contain : \\ / ? * [ ], and not begin or end with an apostrophe.",
                nameof(name));
        }
    }

    private static void EnsureSheetNameAvailable(
        Excel.Workbook workbook, string name, CancellationToken cancellationToken, string? currentName = null)
    {
        Excel.Sheets? sheets = null;
        try
        {
            sheets = workbook.Sheets;
            for (var index = 1; index <= sheets.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                object? candidate = null;
                try
                {
                    candidate = sheets.Item[index];
                    var existingName = candidate switch
                    {
                        Excel.Worksheet worksheet => worksheet.Name,
                        Excel.Chart chart => chart.Name,
                        _ => throw new InvalidOperationException("Excel returned an unsupported sheet type during name validation.")
                    };
                    if (string.Equals(existingName, name, StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(existingName, currentName, StringComparison.Ordinal))
                    {
                        throw new InvalidOperationException($"Sheet '{name}' already exists.");
                    }
                }
                finally { ComUtilities.Release(ref candidate); }
            }
        }
        finally { ComUtilities.Release(ref sheets); }
    }

    /// <inheritdoc />
    public OperationResult Delete(IExcelBatch batch, string sheetName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                }
                sheet.Delete();
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
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
        // Validate parameters
        if (!string.IsNullOrWhiteSpace(beforeSheet) && !string.IsNullOrWhiteSpace(afterSheet))
        {
            throw new ArgumentException("Cannot specify both beforeSheet and afterSheet");
        }

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Worksheet? targetSheet = null;
            dynamic? sheets = null;
            dynamic? lastSheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
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
                        throw new InvalidOperationException($"Target sheet '{targetName}' not found.");
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
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
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

    // === ATOMIC CROSS-FILE OPERATIONS ===

    /// <inheritdoc />
    public OperationResult CopyToFile(string sourceFile, string sourceSheet, string targetFile, string? targetSheetName = null, string? beforeSheet = null, string? afterSheet = null)
    {
        if (targetSheetName is not null)
        {
            ValidateNewSheetName(targetSheetName);
        }

        // Validate positioning parameters
        if (!string.IsNullOrWhiteSpace(beforeSheet) && !string.IsNullOrWhiteSpace(afterSheet))
        {
            throw new ArgumentException("Cannot specify both beforeSheet and afterSheet. Choose one or neither.");
        }

        // Validate file paths
        if (string.IsNullOrWhiteSpace(sourceFile))
            throw new ArgumentException("sourceFile is required", nameof(sourceFile));
        if (string.IsNullOrWhiteSpace(targetFile))
            throw new ArgumentException("targetFile is required", nameof(targetFile));
        if (!File.Exists(sourceFile))
            throw new FileNotFoundException($"Source file not found: {sourceFile}");
        if (!File.Exists(targetFile))
            throw new FileNotFoundException($"Target file not found: {targetFile}");

        // Normalize paths for comparison
        string normalizedSource = Path.GetFullPath(sourceFile);
        string normalizedTarget = Path.GetFullPath(targetFile);
        if (string.Equals(normalizedSource, normalizedTarget, StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("Source and target files must be different. For same-file copy, use the 'copy' action.");
        }

        // Create a batch with both files open in the same Excel instance
        using var batch = ExcelSession.BeginBatch(sourceFile, targetFile);

        return batch.Execute((ctx, ct) =>
        {
            dynamic? sourceWb = null;
            dynamic? targetWb = null;
            Excel.Worksheet? sourceSheetObj = null;
            dynamic? targetSheets = null;
            Excel.Worksheet? targetPositionSheet = null;
            Excel.Worksheet? copiedSheet = null;

            try
            {
                // Get both workbooks from the batch
                sourceWb = batch.GetWorkbook(normalizedSource);
                targetWb = batch.GetWorkbook(normalizedTarget);

                // Find source sheet
                sourceSheetObj = ComUtilities.FindSheet(sourceWb, sourceSheet);
                if (sourceSheetObj == null)
                {
                    throw new InvalidOperationException($"Source sheet '{sourceSheet}' not found in '{Path.GetFileName(sourceFile)}'");
                }

                // Handle positioning
                targetSheets = targetWb.Worksheets;
                int? copiedSheetPosition = null;

                if (targetSheetName is not null)
                {
                    EnsureSheetNameAvailable(
                        (Excel.Workbook)targetWb, targetSheetName, ct);
                }

                if (!string.IsNullOrWhiteSpace(beforeSheet))
                {
                    targetPositionSheet = ComUtilities.FindSheet(targetWb, beforeSheet);
                    if (targetPositionSheet == null)
                    {
                        throw new InvalidOperationException($"Target sheet '{beforeSheet}' not found in '{Path.GetFileName(targetFile)}'");
                    }
                    // Get position before copy - the copied sheet will be at this position
                    copiedSheetPosition = Convert.ToInt32(targetPositionSheet.Index);
                    sourceSheetObj.Copy(Before: targetPositionSheet);
                }
                else if (!string.IsNullOrWhiteSpace(afterSheet))
                {
                    targetPositionSheet = ComUtilities.FindSheet(targetWb, afterSheet);
                    if (targetPositionSheet == null)
                    {
                        throw new InvalidOperationException($"Target sheet '{afterSheet}' not found in '{Path.GetFileName(targetFile)}'");
                    }
                    // Get position before copy - the copied sheet will be at position + 1
                    copiedSheetPosition = Convert.ToInt32(targetPositionSheet.Index) + 1;
                    sourceSheetObj.Copy(After: targetPositionSheet);
                }
                else
                {
                    // Copy to end of target workbook
                    dynamic? lastSheet = targetSheets.Item(targetSheets.Count);
                    try
                    {
                        sourceSheetObj.Copy(After: lastSheet);
                        // Copied sheet will be at the end (new count)
                        copiedSheetPosition = targetSheets.Count;
                    }
                    finally
                    {
                        ComUtilities.Release(ref lastSheet!);
                    }
                }

                // Rename if requested - use correct position based on where sheet was copied
                if (targetSheetName is not null && copiedSheetPosition.HasValue)
                {
                    copiedSheet = (Excel.Worksheet)targetSheets.Item(copiedSheetPosition.Value);
                    var copiedSheetCurrentName = copiedSheet.Name;
                    SetSheetNameWithContext(copiedSheet, targetSheetName, copiedSheetCurrentName, isNewSheet: true);
                }

                // Save the target workbook (source unchanged, only target modified)
                targetWb.Save();

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref copiedSheet);
                ComUtilities.Release(ref targetPositionSheet);
                ComUtilities.Release(ref targetSheets);
                ComUtilities.Release(ref sourceSheetObj);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult MoveToFile(string sourceFile, string sourceSheet, string targetFile, string? beforeSheet = null, string? afterSheet = null)
    {
        // Validate positioning parameters
        if (!string.IsNullOrWhiteSpace(beforeSheet) && !string.IsNullOrWhiteSpace(afterSheet))
        {
            throw new ArgumentException("Cannot specify both beforeSheet and afterSheet. Choose one or neither.");
        }

        // Validate file paths
        if (string.IsNullOrWhiteSpace(sourceFile))
            throw new ArgumentException("sourceFile is required", nameof(sourceFile));
        if (string.IsNullOrWhiteSpace(targetFile))
            throw new ArgumentException("targetFile is required", nameof(targetFile));
        if (!File.Exists(sourceFile))
            throw new FileNotFoundException($"Source file not found: {sourceFile}");
        if (!File.Exists(targetFile))
            throw new FileNotFoundException($"Target file not found: {targetFile}");

        // Normalize paths for comparison
        string normalizedSource = Path.GetFullPath(sourceFile);
        string normalizedTarget = Path.GetFullPath(targetFile);
        if (string.Equals(normalizedSource, normalizedTarget, StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("Source and target files must be different. For same-file move, use the 'move' action.");
        }

        // Create a batch with both files open in the same Excel instance
        using var batch = ExcelSession.BeginBatch(sourceFile, targetFile);

        return batch.Execute((ctx, ct) =>
        {
            dynamic? sourceWb = null;
            dynamic? targetWb = null;
            Excel.Worksheet? sourceSheetObj = null;
            dynamic? targetSheets = null;
            Excel.Worksheet? targetPositionSheet = null;

            try
            {
                // Get both workbooks from the batch
                sourceWb = batch.GetWorkbook(normalizedSource);
                targetWb = batch.GetWorkbook(normalizedTarget);

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
                    sourceSheetObj.Move(Before: targetPositionSheet);
                }
                else if (!string.IsNullOrWhiteSpace(afterSheet))
                {
                    targetPositionSheet = ComUtilities.FindSheet(targetWb, afterSheet);
                    if (targetPositionSheet == null)
                    {
                        throw new InvalidOperationException($"Target sheet '{afterSheet}' not found in '{Path.GetFileName(targetFile)}'");
                    }
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

                // Save both workbooks (source lost a sheet, target gained one)
                sourceWb.Save();
                targetWb.Save();

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref targetPositionSheet);
                ComUtilities.Release(ref targetSheets);
                // Note: sourceSheetObj has been moved, don't release it
            }
        });
    }
}
