using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Sbroenne.ExcelMcp.ComInterop;

namespace Sbroenne.ExcelMcp.Core.Commands.Screenshot;

/// <summary>
/// Works out how to fit a worksheet range into the Excel viewport for capture: which zoom level to
/// use, and how to split the range into tiles when it is still larger than the viewport.
///
/// All worksheet geometry is in points, which are independent of both zoom and display scaling, so
/// the zoom needed to fit a range is simply the ratio of the grid area to the range size. The grid
/// area itself is measured by <see cref="MeasureUsablePane"/> rather than taken from
/// <c>Window.UsableWidth</c>/<c>UsableHeight</c>, which overstate it.
/// </summary>
[SupportedOSPlatform("windows")]
internal static class CapturePlanner
{
    /// <summary>Excel refuses zoom levels below 10%.</summary>
    private const int MinExcelZoom = 10;

    /// <summary>
    /// Preferred lower bound on zoom: below this, cell text stops being legible, so a range that
    /// does not fit is tiled rather than shrunk further. The tile budget wins if the two conflict —
    /// see <see cref="PreferredMaxTiles"/> — so a range that would otherwise need more than the
    /// preferred number of tiles may still be zoomed below this floor, down to its fit zoom.
    /// </summary>
    private const int MinReadableZoom = 40;

    /// <summary>Preferred ceiling on tile count; zoom is reduced to stay within it.</summary>
    private const int PreferredMaxTiles = 36;

    /// <summary>Hard ceiling on tile count. Beyond this the captured range is truncated.</summary>
    private const int HardMaxTiles = 64;

    /// <summary>A contiguous run of rows or columns that fits in one viewport.</summary>
    internal sealed record Segment(int Start, int Count, double Offset, double Size);

    /// <summary>The grid area available for capture, in screen points.</summary>
    internal readonly record struct UsablePane(double Width, double Height);

    /// <summary>The zoom and tile layout to use for a capture.</summary>
    internal sealed record CapturePlan(
        int Zoom,
        IReadOnlyList<Segment> RowSegments,
        IReadOnlyList<Segment> ColumnSegments,
        bool Truncated);

    /// <summary>
    /// Safety margin, in screen points, trimmed off the measured grid area.
    ///
    /// <c>VisibleRange</c> can list a trailing row or column whose leading edge is already past the
    /// painted grid, which is enough to pull a few pixels of chrome into a tile. One default row
    /// height covers that error at every zoom level, and costs at most one extra tile.
    /// </summary>
    private const double GridSafetyPoints = 15.0;

    /// <summary>
    /// Measures the grid area of the pane in screen points, excluding the window chrome around it.
    ///
    /// <c>Window.UsableWidth</c>/<c>UsableHeight</c> describe the whole workspace, which includes
    /// the horizontal scroll bar and sheet tab strip below the grid. Sizing a tile from them reaches
    /// past the last painted grid row, so the capture picks up a strip of chrome at the bottom of
    /// every tile - visible as a band at each seam of a stitched image.
    ///
    /// <c>VisibleRange</c> ends with a row and column that are only partially visible, so measuring
    /// to their leading edge yields an area that is close to the real grid; a further safety margin
    /// covers Excel over-reporting that trailing item.
    /// </summary>
    internal static UsablePane MeasureUsablePane(dynamic window)
    {
        double workspaceWidth = Convert.ToDouble(window.UsableWidth);
        double workspaceHeight = Convert.ToDouble(window.UsableHeight);
        double zoomFactor = Convert.ToDouble(window.Zoom) / 100.0;

        dynamic? visible = null;
        try
        {
            visible = window.VisibleRange;

            return new UsablePane(
                MeasureAxis(visible, false, workspaceWidth, zoomFactor),
                MeasureAxis(visible, true, workspaceHeight, zoomFactor));
        }
        catch (COMException)
        {
            // Excel exposes no visible range for some window states; the workspace size is the only
            // measurement left, so fall back to it rather than failing the capture.
            return new UsablePane(FallbackExtent(workspaceWidth), FallbackExtent(workspaceHeight));
        }
        finally
        {
            ComUtilities.Release(ref visible);
        }
    }

    /// <summary>
    /// Best available extent when the grid cannot be measured from <c>VisibleRange</c>.
    ///
    /// The workspace extent is the only measurement left, but it is exactly the value that reaches
    /// into the chrome, so the safety margin is taken off it as well. That margin comfortably
    /// exceeds the scroll bar and tab strip measured on a standard display, though an unusually tall
    /// chrome could still leak a few pixels in this path.
    /// </summary>
    private static double FallbackExtent(double workspaceExtent)
    {
        double trimmed = workspaceExtent - GridSafetyPoints;

        return trimmed > 0 ? trimmed : workspaceExtent;
    }

    /// <summary>
    /// Measures one axis of the visible grid, in screen points, up to the leading edge of the last
    /// (partially visible) row or column. Falls back to the trimmed workspace extent when the axis
    /// holds a single row or column, which is too large to measure this way.
    /// </summary>
    private static double MeasureAxis(dynamic visible, bool rows, double workspaceExtent, double zoomFactor)
    {
        dynamic? items = null;
        dynamic? first = null;
        dynamic? last = null;

        try
        {
            items = rows ? visible.Rows : visible.Columns;
            int count = Convert.ToInt32(items.Count);

            if (count < 2 || zoomFactor <= 0)
            {
                return FallbackExtent(workspaceExtent);
            }

            first = items[1];
            last = items[count];

            double leading = rows ? Convert.ToDouble(first.Top) : Convert.ToDouble(first.Left);
            double trailing = rows ? Convert.ToDouble(last.Top) : Convert.ToDouble(last.Left);
            double extent = (trailing - leading) * zoomFactor - GridSafetyPoints;

            return extent > 0 ? Math.Min(FallbackExtent(workspaceExtent), extent) : FallbackExtent(workspaceExtent);
        }
        finally
        {
            ComUtilities.Release(ref last);
            ComUtilities.Release(ref first);
            ComUtilities.Release(ref items);
        }
    }

    /// <summary>
    /// Builds the capture plan for a range.
    /// </summary>
    /// <param name="window">Excel window that will perform the capture.</param>
    /// <param name="range">Range to capture.</param>
    public static CapturePlan Plan(dynamic window, dynamic range)
    {
        UsablePane usable = MeasureUsablePane(window);
        double rangeWidth = Convert.ToDouble(range.Width);
        double rangeHeight = Convert.ToDouble(range.Height);

        int fitZoom = CalculateFitZoom(usable.Width, usable.Height, rangeWidth, rangeHeight);
        int zoom = Math.Clamp(fitZoom, Math.Min(MinReadableZoom, 100), 100);

        return BuildPlan(range, usable, zoom, fitZoom);
    }

    /// <summary>
    /// Rebuilds the tile layout for an already-chosen zoom, using the pane as it measures at that
    /// zoom.
    ///
    /// The row and column headers scale with zoom, so the grid area is slightly different once the
    /// capture zoom is applied. Re-deriving the segments from the post-zoom pane keeps the planned
    /// tile size and the capturable tile size identical, which is what stops a hairline gap
    /// appearing at each seam of a stitched image.
    /// </summary>
    /// <param name="range">Range to capture.</param>
    /// <param name="zoom">Zoom percentage the window is currently set to.</param>
    /// <param name="usable">Grid area measured at that zoom.</param>
    public static CapturePlan Replan(dynamic range, int zoom, UsablePane usable)
    {
        return BuildPlan(range, usable, zoom, zoom);
    }

    /// <summary>
    /// Splits the range into tiles for a given pane size, stepping the zoom down while the tile
    /// count exceeds <see cref="PreferredMaxTiles"/> and the fit zoom allows it.
    /// </summary>
    private static CapturePlan BuildPlan(dynamic range, UsablePane usable, int zoom, int fitZoom)
    {
        double usableWidth = usable.Width;
        double usableHeight = usable.Height;
        int rowCount = Convert.ToInt32(range.Rows.Count);
        int columnCount = Convert.ToInt32(range.Columns.Count);

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
