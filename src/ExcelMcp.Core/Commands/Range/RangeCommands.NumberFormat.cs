using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// RangeCommands partial class - Number formatting operations
/// </summary>
public partial class RangeCommands
{
    // === NUMBER FORMAT OPERATIONS ===

    /// <inheritdoc />
    public async Task<RangeNumberFormatResult> GetNumberFormatsAsync(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        var result = new RangeNumberFormatResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sheetName,
            RangeAddress = rangeAddress
        };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress);
                if (range == null)
                {
                    result.Success = false;
                    result.ErrorMessage = RangeHelpers.GetResolveError(sheetName, rangeAddress);
                    return result;
                }

                // Get number formats as 2D array
                object numberFormats = range.NumberFormat;
                
                // Get dimensions
                int rowCount = Convert.ToInt32(range.Rows.Count);
                int columnCount = Convert.ToInt32(range.Columns.Count);

                result.RowCount = rowCount;
                result.ColumnCount = columnCount;

                if (rowCount == 1 && columnCount == 1)
                {
                    // Single cell - numberFormats is a string
                    result.Formats.Add(new List<string> { numberFormats?.ToString() ?? "General" });
                }
                else
                {
                    // Multiple cells - numberFormats is an array
                    object[,] formats = (object[,])numberFormats;
                    
                    for (int row = 1; row <= rowCount; row++)
                    {
                        var rowList = new List<string>();
                        for (int col = 1; col <= columnCount; col++)
                        {
                            var format = formats[row, col]?.ToString() ?? "General";
                            rowList.Add(format);
                        }
                        result.Formats.Add(rowList);
                    }
                }

                result.Success = true;
                result.SuggestedNextActions =
                [
                    "Use 'set-number-format' to apply uniform format",
                    "Use 'set-number-formats' to apply different formats per cell",
                    "See NumberFormatPresets for common format codes"
                ];
                result.WorkflowHint = $"Retrieved number formats for {rowCount}x{columnCount} range";

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Failed to get number formats: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref range);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> SetNumberFormatAsync(IExcelBatch batch, string sheetName, string rangeAddress, string formatCode)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "set-number-format"
        };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress);
                if (range == null)
                {
                    result.Success = false;
                    result.ErrorMessage = RangeHelpers.GetResolveError(sheetName, rangeAddress);
                    return result;
                }

                // Set uniform number format for entire range
                range.NumberFormat = formatCode;

                result.Success = true;
                result.SuggestedNextActions =
                [
                    "Use 'get-values' to see formatted values",
                    "Use 'get-number-formats' to verify format applied",
                    "Consider auto-fitting columns with 'auto-fit-columns'"
                ];
                result.WorkflowHint = $"Applied format '{formatCode}' to range {rangeAddress}";

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Failed to set number format: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref range);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> SetNumberFormatsAsync(IExcelBatch batch, string sheetName, string rangeAddress, List<List<string>> formats)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "set-number-formats"
        };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress);
                if (range == null)
                {
                    result.Success = false;
                    result.ErrorMessage = RangeHelpers.GetResolveError(sheetName, rangeAddress);
                    return result;
                }

                int rowCount = Convert.ToInt32(range.Rows.Count);
                int columnCount = Convert.ToInt32(range.Columns.Count);

                // Validate dimensions match
                if (formats.Count != rowCount)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Format array row count ({formats.Count}) doesn't match range row count ({rowCount})";
                    return result;
                }

                for (int i = 0; i < formats.Count; i++)
                {
                    if (formats[i].Count != columnCount)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Format array row {i + 1} column count ({formats[i].Count}) doesn't match range column count ({columnCount})";
                        return result;
                    }
                }

                // Convert List<List<string>> to 2D array
                object[,] formatArray = new object[rowCount, columnCount];
                for (int row = 0; row < rowCount; row++)
                {
                    for (int col = 0; col < columnCount; col++)
                    {
                        formatArray[row + 1, col + 1] = formats[row][col];
                    }
                }

                // Set number formats cell-by-cell
                range.NumberFormat = formatArray;

                result.Success = true;
                result.SuggestedNextActions =
                [
                    "Use 'get-values' to see formatted values",
                    "Use 'get-number-formats' to verify formats applied",
                    "Consider auto-fitting columns with 'auto-fit-columns'"
                ];
                result.WorkflowHint = $"Applied {rowCount}x{columnCount} number formats to range {rangeAddress}";

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Failed to set number formats: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref range);
            }
        });
    }
}
