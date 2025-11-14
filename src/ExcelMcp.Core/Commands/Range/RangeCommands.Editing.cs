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
    public async Task<OperationResult> ClearAllAsync(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        return await ClearRangeAsync(batch, sheetName, rangeAddress, "clear-all", r => r.Clear());
    }

    /// <inheritdoc />
    public async Task<OperationResult> ClearContentsAsync(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        return await ClearRangeAsync(batch, sheetName, rangeAddress, "clear-contents", r => r.ClearContents());
    }

    /// <inheritdoc />
    public async Task<OperationResult> ClearFormatsAsync(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        return await ClearRangeAsync(batch, sheetName, rangeAddress, "clear-formats", r => r.ClearFormats());
    }

    // === COPY OPERATIONS ===

    /// <inheritdoc />
    public async Task<OperationResult> CopyAsync(IExcelBatch batch, string sourceSheet, string sourceRange, string targetSheet, string targetRange)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "copy" };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? srcRange = null;
            dynamic? tgtRange = null;
            try
            {
                srcRange = RangeHelpers.ResolveRange(ctx.Book, sourceSheet, sourceRange, out string? srcError);
                if (srcRange == null)
                {
                    result.Success = false;
                    result.ErrorMessage = srcError ?? RangeHelpers.GetResolveError(sourceSheet, sourceRange);
                    return result;
                }

                tgtRange = RangeHelpers.ResolveRange(ctx.Book, targetSheet, targetRange, out string? tgtError);
                if (tgtRange == null)
                {
                    result.Success = false;
                    result.ErrorMessage = tgtError ?? RangeHelpers.GetResolveError(targetSheet, targetRange);
                    return result;
                }

                srcRange.Copy(tgtRange);
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
                ComUtilities.Release(ref srcRange);
                ComUtilities.Release(ref tgtRange);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> CopyValuesAsync(IExcelBatch batch, string sourceSheet, string sourceRange, string targetSheet, string targetRange)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "copy-values" };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? srcRange = null;
            dynamic? tgtRange = null;
            try
            {
                srcRange = RangeHelpers.ResolveRange(ctx.Book, sourceSheet, sourceRange, out string? srcError);
                if (srcRange == null)
                {
                    result.Success = false;
                    result.ErrorMessage = srcError ?? RangeHelpers.GetResolveError(sourceSheet, sourceRange);
                    return result;
                }

                tgtRange = RangeHelpers.ResolveRange(ctx.Book, targetSheet, targetRange, out string? tgtError);
                if (tgtRange == null)
                {
                    result.Success = false;
                    result.ErrorMessage = tgtError ?? RangeHelpers.GetResolveError(targetSheet, targetRange);
                    return result;
                }

                srcRange.Copy();
                tgtRange.PasteSpecial(-4163); // xlPasteValues
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
                ComUtilities.Release(ref srcRange);
                ComUtilities.Release(ref tgtRange);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> CopyFormulasAsync(IExcelBatch batch, string sourceSheet, string sourceRange, string targetSheet, string targetRange)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "copy-formulas" };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? srcRange = null;
            dynamic? tgtRange = null;
            try
            {
                srcRange = RangeHelpers.ResolveRange(ctx.Book, sourceSheet, sourceRange, out string? srcError);
                if (srcRange == null)
                {
                    result.Success = false;
                    result.ErrorMessage = srcError ?? RangeHelpers.GetResolveError(sourceSheet, sourceRange);
                    return result;
                }

                tgtRange = RangeHelpers.ResolveRange(ctx.Book, targetSheet, targetRange, out string? tgtError);
                if (tgtRange == null)
                {
                    result.Success = false;
                    result.ErrorMessage = tgtError ?? RangeHelpers.GetResolveError(targetSheet, targetRange);
                    return result;
                }

                srcRange.Copy();
                tgtRange.PasteSpecial(-4123); // xlPasteFormulas
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
                ComUtilities.Release(ref srcRange);
                ComUtilities.Release(ref tgtRange);
            }
        });
    }

    // === INSERT/DELETE OPERATIONS ===

    /// <inheritdoc />
    public async Task<OperationResult> InsertCellsAsync(IExcelBatch batch, string sheetName, string rangeAddress, InsertShiftDirection shift)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "insert-cells" };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    result.Success = false;
                    result.ErrorMessage = specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress);
                    return result;
                }

                int shiftConst = shift == InsertShiftDirection.Down ? -4121 : -4161; // xlShiftDown : xlShiftToRight
                range.Insert(shiftConst);
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
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> DeleteCellsAsync(IExcelBatch batch, string sheetName, string rangeAddress, DeleteShiftDirection shift)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "delete-cells" };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    result.Success = false;
                    result.ErrorMessage = specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress);
                    return result;
                }

                int shiftConst = shift == DeleteShiftDirection.Up ? -4162 : -4159; // xlShiftUp : xlShiftToLeft
                range.Delete(shiftConst);
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
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> InsertRowsAsync(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "insert-rows" };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            dynamic? rows = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    result.Success = false;
                    result.ErrorMessage = specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress);
                    return result;
                }

                rows = range.EntireRow;
                rows.Insert();
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
                ComUtilities.Release(ref rows);
                ComUtilities.Release(ref range);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> DeleteRowsAsync(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "delete-rows" };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            dynamic? rows = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    result.Success = false;
                    result.ErrorMessage = specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress);
                    return result;
                }

                rows = range.EntireRow;
                rows.Delete();
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
                ComUtilities.Release(ref rows);
                ComUtilities.Release(ref range);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> InsertColumnsAsync(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "insert-columns" };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            dynamic? columns = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    result.Success = false;
                    result.ErrorMessage = specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress);
                    return result;
                }

                columns = range.EntireColumn;
                columns.Insert();
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
                ComUtilities.Release(ref columns);
                ComUtilities.Release(ref range);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> DeleteColumnsAsync(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "delete-columns" };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            dynamic? columns = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    result.Success = false;
                    result.ErrorMessage = specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress);
                    return result;
                }

                columns = range.EntireColumn;
                columns.Delete();
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
                ComUtilities.Release(ref columns);
                ComUtilities.Release(ref range);
            }
        });
    }

    // === FIND/REPLACE OPERATIONS ===

}
