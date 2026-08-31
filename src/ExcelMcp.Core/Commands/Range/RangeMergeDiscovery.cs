using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Shared merged-range discovery for range and table safety operations.
/// </summary>
internal static class RangeMergeDiscovery
{
    // Unmerged ranges bypass this scan; mixed ranges require one COM lookup per cell.
    private const long MaxMergedRangeScanCells = 4_096;

    internal static bool? GetMergeCellsState(object? mergeCells)
    {
        if (mergeCells is null || mergeCells == DBNull.Value)
        {
            return null;
        }

        return Convert.ToBoolean(mergeCells, CultureInfo.InvariantCulture);
    }

    internal static List<string> CollectMergedRanges(dynamic range, CancellationToken cancellationToken)
    {
        dynamic? cells = null;
        var mergedRanges = new List<string>();
        var seenRanges = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        try
        {
            long cellCount = Convert.ToInt64(range.CountLarge, CultureInfo.InvariantCulture);
            if (cellCount > MaxMergedRangeScanCells)
            {
                string rangeAddress = Convert.ToString(range.Address, CultureInfo.InvariantCulture)
                    ?? "(unknown range)";
                throw new InvalidOperationException(
                    $"Cannot inspect merged cells in range '{rangeAddress}' because it contains " +
                    $"{cellCount.ToString("N0", CultureInfo.InvariantCulture)} cells, exceeding the safe scan limit " +
                    $"of {MaxMergedRangeScanCells.ToString("N0", CultureInfo.InvariantCulture)} cells. " +
                    "Use a smaller range for this operation, or unmerge the affected cells before retrying.");
            }

            cells = range.Cells;
            int boundedCellCount = checked((int)cellCount);

            for (int i = 1; i <= boundedCellCount; i++)
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
}
