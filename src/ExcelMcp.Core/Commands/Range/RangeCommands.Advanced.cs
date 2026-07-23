using System.Globalization;
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
    public OperationResult MergeCells(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets[sheetName];

                // Get range
                range = sheet.Range[rangeAddress];

                // Merge cells
                range.Merge();

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult UnmergeCells(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets[sheetName];

                // Get range
                range = sheet.Range[rangeAddress];

                // Unmerge cells
                range.UnMerge();

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    /// <inheritdoc />
    public RangeMergeInfoResult GetMergeInfo(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets[sheetName];

                // Get range
                range = sheet.Range[rangeAddress];

                object? mergeCells = range.MergeCells;
                bool? isMergedState = GetMergeCellsState(mergeCells);
                IReadOnlyList<string> mergedRanges = isMergedState == false
                    ? []
                    : CollectMergedRanges(range, ct);

                return new RangeMergeInfoResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    SheetName = sheetName,
                    RangeAddress = rangeAddress,
                    IsMerged = isMergedState ?? mergedRanges.Count > 0,
                    MergedRanges = mergedRanges
                };
            }
            finally
            {
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    private static bool? GetMergeCellsState(object? mergeCells)
    {
        if (mergeCells is null || mergeCells == DBNull.Value)
        {
            return null;
        }

        return Convert.ToBoolean(mergeCells, CultureInfo.InvariantCulture);
    }

    private static List<string> CollectMergedRanges(dynamic range, CancellationToken cancellationToken)
    {
        dynamic? cells = null;
        var mergedRanges = new List<string>();
        var seenRanges = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        try
        {
            cells = range.Cells;
            int cellCount = Convert.ToInt32(cells.Count);

            for (int i = 1; i <= cellCount; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                dynamic? cell = null;
                dynamic? mergeArea = null;
                try
                {
                    cell = cells.Item[i];
                    object? cellMergeCells = cell.MergeCells;

                    if (GetMergeCellsState(cellMergeCells) != true)
                    {
                        continue;
                    }

                    mergeArea = cell.MergeArea;
                    string address = mergeArea.Address?.ToString() ?? string.Empty;
                    if (address.Length > 0 && seenRanges.Add(address))
                    {
                        mergedRanges.Add(address);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref mergeArea);
                    ComUtilities.Release(ref cell);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref cells);
        }

        return mergedRanges;
    }

    /// <inheritdoc />
    public OperationResult SetCellLock(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        bool locked)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets[sheetName];

                // Get range
                range = sheet.Range[rangeAddress];

                // Set locked property
                range.Locked = locked;

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    /// <inheritdoc />
    public RangeLockInfoResult GetCellLock(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets[sheetName];

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
            finally
            {
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

}



