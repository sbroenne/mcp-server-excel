using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Auto-fit operations for Excel ranges (partial class)
/// </summary>
public partial class RangeCommands
{
    /// <inheritdoc />
    public async Task<OperationResult> AutoFitColumnsAsync(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        return await batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? columns = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets.Item(sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Get columns and auto-fit
                columns = range.Columns;
                columns.AutoFit();

                return new OperationResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath
                };
            }
            catch (Exception ex)
            {
                return new OperationResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to auto-fit columns for range '{rangeAddress}': {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref columns!);
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> AutoFitRowsAsync(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        return await batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? rows = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets.Item(sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Get rows and auto-fit
                rows = range.Rows;
                rows.AutoFit();

                return new OperationResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath
                };
            }
            catch (Exception ex)
            {
                return new OperationResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to auto-fit rows for range '{rangeAddress}': {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref rows!);
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }
}
