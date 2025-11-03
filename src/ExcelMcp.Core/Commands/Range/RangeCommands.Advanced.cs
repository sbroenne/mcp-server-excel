using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Merge, conditional formatting, and protection operations for Excel ranges (partial class)
/// </summary>
public partial class RangeCommands
{
    /// <inheritdoc />
    public async Task<OperationResult> MergeCellsAsync(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        return await batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets.Item(sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Merge cells
                range.Merge();

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
                    ErrorMessage = $"Failed to merge cells in range '{rangeAddress}': {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> UnmergeCellsAsync(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        return await batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets.Item(sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Unmerge cells
                range.UnMerge();

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
                    ErrorMessage = $"Failed to unmerge cells in range '{rangeAddress}': {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    /// <inheritdoc />
    public async Task<RangeMergeInfoResult> GetMergeInfoAsync(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        return await batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets.Item(sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Check if merged
                var isMerged = range.MergeCells ?? false;

                return new RangeMergeInfoResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    SheetName = sheetName,
                    RangeAddress = rangeAddress,
                    IsMerged = isMerged
                };
            }
            catch (Exception ex)
            {
                return new RangeMergeInfoResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to get merge info for range '{rangeAddress}': {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> AddConditionalFormattingAsync(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        string ruleType,
        string? formula1,
        string? formula2,
        string? formatStyle)
    {
        return await batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? formatConditions = null;
            dynamic? formatCondition = null;
            dynamic? interior = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets.Item(sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Get format conditions
                formatConditions = range.FormatConditions;

                // Parse rule type
                var xlType = ParseConditionalFormattingType(ruleType);

                // Add format condition
                formatCondition = formatConditions.Add(
                    Type: xlType,
                    Operator: 3, // xlBetween (default, can be parameterized later)
                    Formula1: formula1 ?? "",
                    Formula2: formula2 ?? "");

                // Apply format style if specified
                if (!string.IsNullOrEmpty(formatStyle))
                {
                    interior = formatCondition.Interior;
                    interior.Color = ParseColor(formatStyle);
                }

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
                    ErrorMessage = $"Failed to add conditional formatting to range '{rangeAddress}': {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref interior!);
                ComUtilities.Release(ref formatCondition!);
                ComUtilities.Release(ref formatConditions!);
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> ClearConditionalFormattingAsync(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        return await batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? formatConditions = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets.Item(sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Get format conditions and delete all
                formatConditions = range.FormatConditions;
                formatConditions.Delete();

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
                    ErrorMessage = $"Failed to clear conditional formatting from range '{rangeAddress}': {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref formatConditions!);
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> SetCellLockAsync(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        bool locked)
    {
        return await batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets.Item(sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Set locked property
                range.Locked = locked;

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
                    ErrorMessage = $"Failed to set cell lock for range '{rangeAddress}': {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    /// <inheritdoc />
    public async Task<RangeLockInfoResult> GetCellLockAsync(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        return await batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets.Item(sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Get locked property
                var isLocked = range.Locked ?? false;

                return new RangeLockInfoResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    SheetName = sheetName,
                    RangeAddress = rangeAddress,
                    IsLocked = isLocked
                };
            }
            catch (Exception ex)
            {
                return new RangeLockInfoResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to get cell lock info for range '{rangeAddress}': {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    private static int ParseConditionalFormattingType(string type)
    {
        return type.ToLowerInvariant() switch
        {
            "cellvalue" => 1, // xlCellValue
            "expression" => 2, // xlExpression
            "colorscale" => 3, // xlColorScale
            "databar" => 4, // xlDatabar
            "top10" => 5, // xlTop10
            "iconset" => 6, // xlIconSet
            "uniquevalues" => 8, // xlUniqueValues
            "blanksCondition" => 10, // xlBlanksCondition
            "timePeriod" => 11, // xlTimePeriod
            "aboveaverage" => 12, // xlAboveAverageCondition
            _ => throw new ArgumentException($"Invalid conditional formatting type: {type}")
        };
    }
}
