using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;


namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Range editing operations (clear, copy, insert/delete cells/rows/columns)
/// </summary>
public partial class RangeCommands
{
    // === CLEAR OPERATIONS ===

    /// <summary>
    /// Clears all content (values, formulas, formats) from range
    /// Excel COM: Range.Clear()
    /// </summary>
    public OperationResult ClearAll(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        return ClearRange(batch, sheetName, rangeAddress, "clear-all", r => r.Clear());
    }

    /// <inheritdoc />
    public OperationResult ClearContents(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        return ClearRange(batch, sheetName, rangeAddress, "clear-contents", r => r.ClearContents());
    }

    /// <inheritdoc />
    public OperationResult ClearFormats(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        return ClearRange(batch, sheetName, rangeAddress, "clear-formats", r => r.ClearFormats());
    }

    // === COPY OPERATIONS ===

    /// <inheritdoc />
    public OperationResult Copy(IExcelBatch batch, string sourceSheet, string sourceRange, string targetSheet, string targetRange)
    {
        return CopyRange(batch, sourceSheet, sourceRange, targetSheet, targetRange, "copy",
            (src, tgt) => src.Copy(tgt));
    }

    /// <inheritdoc />
    public OperationResult CopyValues(IExcelBatch batch, string sourceSheet, string sourceRange, string targetSheet, string targetRange)
    {
        return CopyRange(batch, sourceSheet, sourceRange, targetSheet, targetRange, "copy-values",
            (src, tgt) => { src.Copy(); tgt.PasteSpecial(-4163); }); // xlPasteValues
    }

    /// <inheritdoc />
    public OperationResult CopyFormulas(IExcelBatch batch, string sourceSheet, string sourceRange, string targetSheet, string targetRange)
    {
        return CopyRange(batch, sourceSheet, sourceRange, targetSheet, targetRange, "copy-formulas",
            (src, tgt) => { src.Copy(); tgt.PasteSpecial(-4123); }); // xlPasteFormulas
    }

    // === INSERT/DELETE OPERATIONS ===

    /// <inheritdoc />
    public OperationResult InsertCells(IExcelBatch batch, string sheetName, string rangeAddress, InsertShiftDirection insertShift)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "insert-cells" };

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

                int shiftConst = insertShift == InsertShiftDirection.Down ? -4121 : -4161; // xlShiftDown : xlShiftToRight
                range.Insert(shiftConst);
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref range);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult DeleteCells(IExcelBatch batch, string sheetName, string rangeAddress, DeleteShiftDirection deleteShift)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "delete-cells" };

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

                int shiftConst = deleteShift == DeleteShiftDirection.Up ? -4162 : -4159; // xlShiftUp : xlShiftToLeft
                range.Delete(shiftConst);
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref range);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult InsertRows(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        return ModifyRowsOrColumns(batch, sheetName, rangeAddress, "insert-rows",
            r => r.EntireRow, rows => rows.Insert());
    }

    /// <inheritdoc />
    public OperationResult DeleteRows(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        return ModifyRowsOrColumns(batch, sheetName, rangeAddress, "delete-rows",
            r => r.EntireRow, rows => rows.Delete());
    }

    /// <inheritdoc />
    public OperationResult InsertColumns(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        return ModifyRowsOrColumns(batch, sheetName, rangeAddress, "insert-columns",
            r => r.EntireColumn, cols => cols.Insert());
    }

    /// <inheritdoc />
    public OperationResult DeleteColumns(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        return ModifyRowsOrColumns(batch, sheetName, rangeAddress, "delete-columns",
            r => r.EntireColumn, cols => cols.Delete());
    }

    // === FIND/REPLACE OPERATIONS ===

}



