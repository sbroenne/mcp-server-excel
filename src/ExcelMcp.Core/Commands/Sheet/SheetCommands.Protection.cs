using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet protection operations.
/// </summary>
public partial class SheetCommands
{
    /// <inheritdoc />
    public OperationResult SetProtection(IExcelBatch batch, string sheetName, bool isProtected,
        string? password = null, SheetProtectionOptions? options = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        if (!isProtected && options is not null)
            throw new ArgumentException("Protection options are valid only when protecting.", nameof(options));
        if (options?.Selection is { } selection && !Enum.IsDefined(selection))
            throw new ArgumentOutOfRangeException(nameof(options), "Unknown protected-sheet selection mode.");
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            try
            {
                ct.ThrowIfCancellationRequested();
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                }

                if (isProtected)
                {
                    var settings = options ?? new SheetProtectionOptions();
                    sheet.Protect(Password: password is null ? Type.Missing : password,
                        DrawingObjects: settings.DrawingObjects, Contents: settings.Contents,
                        Scenarios: settings.Scenarios, UserInterfaceOnly: settings.UserInterfaceOnly,
                        AllowFormattingCells: settings.AllowFormattingCells,
                        AllowFormattingColumns: settings.AllowFormattingColumns,
                        AllowFormattingRows: settings.AllowFormattingRows,
                        AllowInsertingColumns: settings.AllowInsertingColumns,
                        AllowInsertingRows: settings.AllowInsertingRows,
                        AllowInsertingHyperlinks: settings.AllowInsertingHyperlinks,
                        AllowDeletingColumns: settings.AllowDeletingColumns,
                        AllowDeletingRows: settings.AllowDeletingRows,
                        AllowSorting: settings.AllowSorting, AllowFiltering: settings.AllowFiltering,
                        AllowUsingPivotTables: settings.AllowUsingPivotTables);
                    if (settings.Selection.HasValue)
                        sheet.EnableSelection = settings.Selection.Value switch
                        {
                            ProtectedSheetSelection.AllCells => Excel.XlEnableSelection.xlNoRestrictions,
                            ProtectedSheetSelection.UnlockedCells => Excel.XlEnableSelection.xlUnlockedCells,
                            ProtectedSheetSelection.None => Excel.XlEnableSelection.xlNoSelection,
                            _ => throw new ArgumentOutOfRangeException(nameof(options))
                        };
                }
                else
                {
                    sheet.Unprotect(password is null ? Type.Missing : password);
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public SheetProtectionResult GetProtection(IExcelBatch batch, string sheetName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        var result = new SheetProtectionResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Protection? protection = null;
            try
            {
                ct.ThrowIfCancellationRequested();
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                }

                protection = sheet.Protection;
                result.ProtectContents = sheet.ProtectContents;
                result.ProtectDrawingObjects = sheet.ProtectDrawingObjects;
                result.ProtectScenarios = sheet.ProtectScenarios;
                result.IsProtected = result.ProtectContents || result.ProtectDrawingObjects || result.ProtectScenarios;
                result.UserInterfaceOnly = sheet.ProtectionMode;
                result.Permissions = new SheetProtectionOptions
                {
                    DrawingObjects = result.ProtectDrawingObjects,
                    Contents = result.ProtectContents,
                    Scenarios = result.ProtectScenarios,
                    UserInterfaceOnly = result.UserInterfaceOnly,
                    AllowFormattingCells = protection.AllowFormattingCells,
                    AllowFormattingColumns = protection.AllowFormattingColumns,
                    AllowFormattingRows = protection.AllowFormattingRows,
                    AllowInsertingColumns = protection.AllowInsertingColumns,
                    AllowInsertingRows = protection.AllowInsertingRows,
                    AllowInsertingHyperlinks = protection.AllowInsertingHyperlinks,
                    AllowDeletingColumns = protection.AllowDeletingColumns,
                    AllowDeletingRows = protection.AllowDeletingRows,
                    AllowSorting = protection.AllowSorting,
                    AllowFiltering = protection.AllowFiltering,
                    AllowUsingPivotTables = protection.AllowUsingPivotTables,
                    Selection = sheet.EnableSelection switch
                    {
                        Excel.XlEnableSelection.xlNoRestrictions => ProtectedSheetSelection.AllCells,
                        Excel.XlEnableSelection.xlUnlockedCells => ProtectedSheetSelection.UnlockedCells,
                        Excel.XlEnableSelection.xlNoSelection => ProtectedSheetSelection.None,
                        _ => throw new InvalidOperationException("Excel returned an unknown protected-sheet selection mode.")
                    }
                };
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref protection);
                ComUtilities.Release(ref sheet);
            }
        });
    }
}
