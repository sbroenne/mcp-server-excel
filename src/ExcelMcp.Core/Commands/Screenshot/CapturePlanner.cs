using System.Runtime.Versioning;
using Sbroenne.ExcelMcp.ComInterop;

namespace Sbroenne.ExcelMcp.Core.Commands.Screenshot;

/// <summary>
/// Works out how to fit a worksheet range into the Excel viewport for capture: which zoom level to
/// use, and how to split the range into tiles when it is still larger than the viewport.
///
/// All worksheet geometry is in points, which are independent of both zoom and display scaling.
/// <c>Window.UsableWidth</c>/<c>UsableHeight</c> describe the pane in screen points, so the zoom
/// needed to fit a range is simply the ratio of the two.
/// </summary>
[SupportedOSPlatform("windows")]
internal static class CapturePlanner
{
    /// <summary>Excel refuses zoom levels below 10%.</summary>
    private const int MinExcelZoom = 10;

    /// <summary>Below this zoom, cell text stops being legible, so tiling is preferred instead.</summary>
    private const int MinReadableZoom = 40;

    /// <summary>Preferred ceiling on tile count; zoom is reduced to stay within it.</summary>
    private const int PreferredMaxTiles = 36;

    /// <summary>Hard ceiling on tile count. Beyond this the captured range is truncated.</summary>
    private const int HardMaxTiles = 64;

    /// <summary>A contiguous run of rows or columns that fits in one viewport.</summary>
    internal sealed record Segment(int Start, int Count, double Offset, double Size);

    /// <summary>The zoom and tile layout to use for a capture.</summary>
    internal sealed record CapturePlan(
        int Zoom,
        IReadOnlyList<Segment> RowSegments,
        IReadOnlyList<Segment> ColumnSegments,
        bool Truncated);

    /// <summary>
    /// Builds the capture plan for a range.
    /// </summary>
    /// <param name="window">Excel window that will perform the capture.</param>
    /// <param name="range">Range to capture.</param>
    public static CapturePlan Plan(dynamic window, dynamic range)
    {
        double usableWidth = Convert.ToDouble(window.UsableWidth);
        double usableHeight = Convert.ToDouble(window.UsableHeight);
        double rangeWidth = Convert.ToDouble(range.Width);
        double rangeHeight = Convert.ToDouble(range.Height);
        int rowCount = Convert.ToInt32(range.Rows.Count);
        int columnCount = Convert.ToInt32(range.Columns.Count);

        int fitZoom = CalculateFitZoom(usableWidth, usableHeight, rangeWidth, rangeHeight);
        int zoom = Math.Clamp(fitZoom, Math.Min(MinReadableZoom, 100), 100);

        List<Segment> rowSegments;
        List<Segment> columnSegments;

        while (true)
        {
            columnSegments = BuildSegments(range, false, columnCount, usableWidth * 100.0 / zoom);
            rowSegments = BuildSegments(range, true, rowCount, usableHeight * 100.0 / zoom);

            int tiles = rowSegments.Count * columnSegments.Count;

            if (tiles <= PreferredMaxTiles || zoom <= fitZoom)
            {
                break;
            }

            zoom = Math.Max(fitZoom, zoom - 5);
        }

        bool truncated = false;

        if (rowSegments.Count * columnSegments.Count > HardMaxTiles)
        {
            // Keep capture time bounded. Prefer keeping the top-left of the range, which is where
            // the caller's content starts, and say so in the result message.
            truncated = true;
            int maxPerAxis = Math.Max(1, (int)Math.Sqrt(HardMaxTiles));
            columnSegments = [.. columnSegments.Take(Math.Min(columnSegments.Count, maxPerAxis))];
            rowSegments = [.. rowSegments.Take(Math.Min(rowSegments.Count, maxPerAxis))];
        }

        return new CapturePlan(zoom, rowSegments, columnSegments, truncated);
    }

    /// <summary>
    /// Calculates the zoom percentage at which the range exactly fits the viewport.
    /// </summary>
    private static int CalculateFitZoom(double usableWidth, double usableHeight, double rangeWidth, double rangeHeight)
    {
        if (rangeWidth <= 0 || rangeHeight <= 0 || usableWidth <= 0 || usableHeight <= 0)
        {
            return 100;
        }

        double fit = Math.Min(usableWidth / rangeWidth, usableHeight / rangeHeight) * 100.0;

        return Math.Clamp((int)Math.Floor(fit), MinExcelZoom, 100);
    }

    /// <summary>
    /// Splits the range's rows or columns into runs that each fit within one viewport.
    /// </summary>
    /// <param name="range">Range being captured.</param>
    /// <param name="rows">True to segment rows, false to segment columns.</param>
    /// <param name="count">Number of rows or columns in the range.</param>
    /// <param name="tilePoints">Viewport size in worksheet points at the planned zoom.</param>
    private static List<Segment> BuildSegments(dynamic range, bool rows, int count, double tilePoints)
    {
        var segments = new List<Segment>();

        if (count <= 0)
        {
            return segments;
        }

        double origin = GetOffset(range, rows, 1);
        int start = 1;

        while (start <= count)
        {
            double startOffset = GetOffset(range, rows, start);
            int end = FindSegmentEnd(range, rows, start, count, startOffset, tilePoints);
            double endOffset = GetOffset(range, rows, end) + GetExtent(range, rows, end);

            segments.Add(new Segment(start, end - start + 1, startOffset - origin, endOffset - startOffset));
            start = end + 1;
        }

        return segments;
    }

    /// <summary>
    /// Finds the last row/column index that still fits within one viewport, via binary search over
    /// the monotonically increasing cell offsets. Always returns at least the starting index, so a
    /// single cell larger than the viewport still makes progress.
    /// </summary>
    private static int FindSegmentEnd(dynamic range, bool rows, int start, int count, double startOffset, double tilePoints)
    {
        int low = start;
        int high = count;
        int best = start;

        while (low <= high)
        {
            int mid = low + ((high - low) / 2);
            double extent = GetOffset(range, rows, mid) + GetExtent(range, rows, mid) - startOffset;

            if (extent <= tilePoints)
            {
                best = mid;
                low = mid + 1;
            }
            else
            {
                high = mid - 1;
            }
        }

        return best;
    }

    /// <summary>Gets the offset in points of a row/column within the sheet.</summary>
    private static double GetOffset(dynamic range, bool rows, int index)
    {
        dynamic? item = null;
        try
        {
            item = rows ? range.Rows[index] : range.Columns[index];
            return rows ? Convert.ToDouble(item.Top) : Convert.ToDouble(item.Left);
        }
        finally
        {
            ComUtilities.Release(ref item);
        }
    }

    /// <summary>Gets the height/width in points of a row/column.</summary>
    private static double GetExtent(dynamic range, bool rows, int index)
    {
        dynamic? item = null;
        try
        {
            item = rows ? range.Rows[index] : range.Columns[index];
            return rows ? Convert.ToDouble(item.Height) : Convert.ToDouble(item.Width);
        }
        finally
        {
            ComUtilities.Release(ref item);
        }
    }
}
