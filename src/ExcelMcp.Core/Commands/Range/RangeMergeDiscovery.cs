using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Shared merged-range discovery for range and table safety operations.
/// </summary>
internal static class RangeMergeDiscovery
{
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
}
