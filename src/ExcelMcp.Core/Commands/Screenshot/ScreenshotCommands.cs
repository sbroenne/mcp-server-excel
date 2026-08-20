using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Core.Commands.Screenshot;

/// <summary>
/// Implementation of screenshot commands.
///
/// Captures the real Excel window with the Win32 <c>PrintWindow</c> API and crops to the requested
/// range. Nothing is written to the workbook and the clipboard is never touched, so capture works
/// on protected worksheets and leaves no trace in the file (issue #777).
/// </summary>
public class ScreenshotCommands : IScreenshotCommands
{
    // Excel WindowState constants
    private const int XlNormal = -4143;
    private const int XlMinimized = -4140;
    private const int SwRestore = 9;

    // Excel needs a moment to render after the window becomes visible or the view changes.
    private const int VisibilitySettleMs = 600;
    private const int WindowStateSettleMs = 300;
    private const int ViewSettleMs = 300;

    /// <summary>
    /// Captures a specific range as an image.
    /// </summary>
    public ScreenshotResult CaptureRange(IExcelBatch batch, string? sheetName = null, string rangeAddress = "A1:Z30", ScreenshotQuality quality = ScreenshotQuality.Medium)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            try
            {
                sheet = string.IsNullOrWhiteSpace(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets[sheetName];

                range = sheet.Range[rangeAddress];
                string actualSheet = sheet.Name?.ToString() ?? "Sheet1";
                string actualRange = range.Address?.ToString() ?? rangeAddress;

                return ExportRangeAsImage(ctx.App, sheet, range, actualSheet, actualRange, quality);
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <summary>
    /// Captures the entire used area of a worksheet as an image.
    /// If UsedRange exceeds 500 rows or 50 columns, it is capped to keep the capture legible
    /// on sheets with formatting extending far beyond the data.
    /// </summary>
    public ScreenshotResult CaptureSheet(IExcelBatch batch, string? sheetName = null, ScreenshotQuality quality = ScreenshotQuality.Medium)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? usedRange = null;
            dynamic? captureRange = null;
            try
            {
                sheet = string.IsNullOrWhiteSpace(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets[sheetName];

                usedRange = sheet.UsedRange;
                string actualSheet = sheet.Name?.ToString() ?? "Sheet1";

                int rows = Convert.ToInt32(usedRange.Rows.Count);
                int cols = Convert.ToInt32(usedRange.Columns.Count);

                const int maxRows = 500;
                const int maxCols = 50;

                if (rows > maxRows || cols > maxCols)
                {
                    int startRow = Convert.ToInt32(usedRange.Row);
                    int startCol = Convert.ToInt32(usedRange.Column);
                    int endRow = startRow + Math.Min(rows, maxRows) - 1;
                    int endCol = startCol + Math.Min(cols, maxCols) - 1;
                    captureRange = sheet.Range[sheet.Cells[startRow, startCol], sheet.Cells[endRow, endCol]];
                }

                dynamic rangeToCapture = captureRange ?? usedRange;
                string actualRange = rangeToCapture.Address?.ToString() ?? "A1";

                return ExportRangeAsImage(ctx.App, sheet, rangeToCapture, actualSheet, actualRange, quality);
            }
            finally
            {
                ComUtilities.Release(ref captureRange);
                ComUtilities.Release(ref usedRange);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <summary>
    /// Captures a range by photographing the Excel window and cropping to the range rectangle.
    ///
    /// The window must be visible for Windows to render it, so a hidden or minimized Excel is
    /// temporarily shown. Zoom and scroll position are adjusted to bring the range into view and
    /// are restored afterwards.
    /// </summary>
    private static ScreenshotResult ExportRangeAsImage(dynamic app, dynamic sheet, dynamic range, string sheetName, string rangeAddress, ScreenshotQuality quality)
    {
        WindowCapture.EnsureDpiAwareness();

        dynamic? window = null;
        dynamic? previousSheet = null;
        Bitmap? composed = null;

        bool visibilityChanged = false;
        bool screenUpdatingChanged = false;
        int originalWindowState = XlNormal;
        bool windowStateChanged = false;
        int originalZoom = 100;
        bool zoomChanged = false;
        int originalScrollRow = 1;
        int originalScrollColumn = 1;
        bool scrollChanged = false;

        try
        {
            bool wasVisible = (bool)app.Visible;
            if (!wasVisible)
            {
                app.Visible = true;
                visibilityChanged = true;
                Thread.Sleep(VisibilitySettleMs);
            }

            // Batch execution suppresses ScreenUpdating for speed, which leaves Excel showing its
            // last painted frame. A screenshot is the one operation that needs Excel to actually
            // redraw, so turn painting back on for the duration of the capture.
            if (!(bool)app.ScreenUpdating)
            {
                app.ScreenUpdating = true;
                screenUpdatingChanged = true;
            }

            originalWindowState = Convert.ToInt32(app.WindowState);
            if (originalWindowState == XlMinimized)
            {
                app.WindowState = XlNormal;
                windowStateChanged = true;
                Thread.Sleep(WindowStateSettleMs);
            }

            previousSheet = GetActiveSheet(app);
            sheet.Activate();

            window = app.ActiveWindow;
            if (window == null)
            {
                throw new InvalidOperationException(
                    "Excel has no active window to capture. Open the workbook in a visible window and retry the screenshot.");
            }

            BringExcelToForeground(app);

            originalZoom = Convert.ToInt32(window.Zoom);
            originalScrollRow = Convert.ToInt32(window.ScrollRow);
            originalScrollColumn = Convert.ToInt32(window.ScrollColumn);

            CapturePlanner.CapturePlan plan = CapturePlanner.Plan(window, range);

            if (plan.Zoom != originalZoom)
            {
                window.Zoom = plan.Zoom;
                zoomChanged = true;
                Thread.Sleep(ViewSettleMs);
            }

            scrollChanged = true;
            composed = CaptureTiles(app, window, range, ref plan);

            var encoded = ScreenshotEncoder.Encode(composed, quality);

            string message = $"Captured {rangeAddress} on '{sheetName}' ({encoded.Width}x{encoded.Height}px)";

            if (plan.RowSegments.Count * plan.ColumnSegments.Count > 1)
            {
                message += $" from {plan.RowSegments.Count * plan.ColumnSegments.Count} tiles at {plan.Zoom}% zoom";
            }

            if (plan.Truncated)
            {
                message += ". The range was too large to capture in full and was truncated to its top-left portion - capture a smaller range for complete output";
            }

            return new ScreenshotResult
            {
                Success = true,
                ImageBase64 = Convert.ToBase64String(encoded.Data),
                MimeType = encoded.MimeType,
                Width = encoded.Width,
                Height = encoded.Height,
                SheetName = sheetName,
                RangeAddress = rangeAddress,
                Message = message
            };
        }
        finally
        {
            if (window != null)
            {
                // Zoom first: Excel scrolls to keep the selection visible when zoom changes, so
                // restoring the scroll position afterwards is what actually sticks.
                if (zoomChanged)
                {
                    try { window.Zoom = originalZoom; } catch (COMException) { }
                }

                if (scrollChanged)
                {
                    try { window.ScrollRow = originalScrollRow; } catch (COMException) { }
                    try { window.ScrollColumn = originalScrollColumn; } catch (COMException) { }
                }
            }

            if (previousSheet != null)
            {
                try { previousSheet.Activate(); } catch (COMException) { }
            }

            if (windowStateChanged)
            {
                try { app.WindowState = originalWindowState; } catch (COMException) { }
            }

            if (screenUpdatingChanged)
            {
                try { app.ScreenUpdating = false; } catch (COMException) { }
            }

            if (visibilityChanged)
            {
                try { app.Visible = false; } catch (COMException) { }
            }

            composed?.Dispose();
            ComUtilities.Release(ref previousSheet);
            ComUtilities.Release(ref window);
        }
    }

    /// <summary>
    /// Captures each planned tile and stitches them into a single bitmap.
    /// </summary>
    private static Bitmap CaptureTiles(dynamic app, dynamic window, dynamic range, ref CapturePlanner.CapturePlan plan)
    {
        IntPtr hwnd = GetExcelWindowHandle(app);
        int dpi = WindowCapture.GetWindowDpi(hwnd);
        double deviceScale = dpi / 72.0;
        double pixelsPerPoint = deviceScale * plan.Zoom / 100.0;

        PaneOrigin paneOrigin = MeasurePaneOrigin(window);

        // Measured after the capture zoom is applied and at the settled A1 scroll position, so that
        // the planned tile size and the capturable tile size agree exactly.
        CapturePlanner.UsablePane usable = CapturePlanner.MeasureUsablePane(window);
        int paneMaxWidth = (int)Math.Round(usable.Width * deviceScale);
        int paneMaxHeight = (int)Math.Round(usable.Height * deviceScale);

        plan = CapturePlanner.Replan(range, plan.Zoom, usable);

        int[] columnOffsets = BuildPixelOffsets(plan.ColumnSegments, pixelsPerPoint, paneMaxWidth);
        int[] rowOffsets = BuildPixelOffsets(plan.RowSegments, pixelsPerPoint, paneMaxHeight);

        int totalWidth = Math.Max(1, columnOffsets[^1]);
        int totalHeight = Math.Max(1, rowOffsets[^1]);

        var canvas = new Bitmap(totalWidth, totalHeight, PixelFormat.Format32bppArgb);
        dynamic? sheet = range.Worksheet;

        try
        {
            using var graphics = Graphics.FromImage(canvas);
            graphics.Clear(Color.White);

            bool verifiedNotBlank = false;

            for (int rowIndex = 0; rowIndex < plan.RowSegments.Count; rowIndex++)
            {
                for (int columnIndex = 0; columnIndex < plan.ColumnSegments.Count; columnIndex++)
                {
                    CapturePlanner.Segment rowSegment = plan.RowSegments[rowIndex];
                    CapturePlanner.Segment columnSegment = plan.ColumnSegments[columnIndex];

                    dynamic? tileRange = null;
                    Bitmap? shot = null;

                    try
                    {
                        tileRange = GetTileRange(sheet, range, rowSegment, columnSegment);

                        ScrollIntoView(window, tileRange);

                        Rectangle windowBounds = WindowCapture.GetWindowBounds(hwnd);
                        shot = WindowCapture.CaptureWindow(hwnd);

                        if (!verifiedNotBlank)
                        {
                            EnsureWindowRendered(shot);
                            verifiedNotBlank = true;
                        }

                        // Excel clamps scrolling near the sheet edges, so measure how far the tile
                        // actually sits from the pane corner rather than assuming it landed there.
                        int offsetX = (int)Math.Round((Convert.ToDouble(tileRange.Left) - GetColumnLeft(sheet, Convert.ToInt32(window.ScrollColumn))) * pixelsPerPoint);
                        int offsetY = (int)Math.Round((Convert.ToDouble(tileRange.Top) - GetRowTop(sheet, Convert.ToInt32(window.ScrollRow))) * pixelsPerPoint);

                        int originX = (int)Math.Round(paneOrigin.X) + offsetX;
                        int originY = (int)Math.Round(paneOrigin.Y) + offsetY;

                        int tileWidth = Math.Min(
                            columnOffsets[columnIndex + 1] - columnOffsets[columnIndex],
                            paneMaxWidth - offsetX);
                        int tileHeight = Math.Min(
                            rowOffsets[rowIndex + 1] - rowOffsets[rowIndex],
                            paneMaxHeight - offsetY);

                        var source = new Rectangle(
                            originX - windowBounds.X,
                            originY - windowBounds.Y,
                            Math.Max(1, tileWidth),
                            Math.Max(1, tileHeight));

                        var destination = new Point(columnOffsets[columnIndex], rowOffsets[rowIndex]);

                        DrawTile(graphics, shot, source, destination);
                    }
                    finally
                    {
                        shot?.Dispose();
                        ComUtilities.Release(ref tileRange);
                    }
                }
            }

            return canvas;
        }
        catch
        {
            canvas.Dispose();
            throw;
        }
        finally
        {
            ComUtilities.Release(ref sheet);
        }
    }

    /// <summary>
    /// Screen pixel position of the grid's top-left corner, below the ribbon and headers.
    /// </summary>
    private readonly record struct PaneOrigin(double X, double Y);

    /// <summary>
    /// Measures where the grid starts on screen by scrolling to cell A1, where the pane corner and
    /// the sheet origin coincide.
    ///
    /// <c>PointsToScreenPixelsX/Y</c> mixes the scroll offset into its result in a way that does not
    /// survive a zoom change, so it is only trustworthy at an unscrolled position. Row and column
    /// headers scale with zoom, so this must be measured after the capture zoom is applied.
    /// </summary>
    private static PaneOrigin MeasurePaneOrigin(dynamic window)
    {
        try
        {
            window.ScrollRow = 1;
            window.ScrollColumn = 1;
        }
        catch (COMException)
        {
            // Fall through: the values below still describe wherever Excel actually is.
        }

        Thread.Sleep(ViewSettleMs);

        return new PaneOrigin(
            Convert.ToDouble(window.PointsToScreenPixelsX(0)),
            Convert.ToDouble(window.PointsToScreenPixelsY(0)));
    }

    /// <summary>Gets the worksheet point offset of a row's top edge.</summary>
    private static double GetRowTop(dynamic sheet, int row)
    {
        dynamic? cell = null;
        try
        {
            cell = sheet.Cells[row, 1];
            return Convert.ToDouble(cell.Top);
        }
        finally
        {
            ComUtilities.Release(ref cell);
        }
    }

    /// <summary>Gets the worksheet point offset of a column's left edge.</summary>
    private static double GetColumnLeft(dynamic sheet, int column)
    {
        dynamic? cell = null;
        try
        {
            cell = sheet.Cells[1, column];
            return Convert.ToDouble(cell.Left);
        }
        finally
        {
            ComUtilities.Release(ref cell);
        }
    }

    /// <summary>
    /// Copies the visible part of a tile from the window capture onto the canvas, clipping to what
    /// the window actually contains so an off-screen edge cannot stretch or wrap the image.
    /// </summary>
    private static void DrawTile(Graphics graphics, Bitmap shot, Rectangle source, Point destination)
    {
        Rectangle clipped = Rectangle.Intersect(source, new Rectangle(0, 0, shot.Width, shot.Height));

        if (clipped.Width <= 0 || clipped.Height <= 0)
        {
            return;
        }

        var target = new Rectangle(
            destination.X + (clipped.X - source.X),
            destination.Y + (clipped.Y - source.Y),
            clipped.Width,
            clipped.Height);

        graphics.DrawImage(shot, target, clipped, GraphicsUnit.Pixel);
    }

    /// <summary>
    /// Converts segment sizes in points to cumulative pixel offsets, so adjacent tiles meet without
    /// gaps or overlaps after rounding.
    /// </summary>
    private static int[] BuildPixelOffsets(IReadOnlyList<CapturePlanner.Segment> segments, double pixelsPerPoint, int paneMax)
    {
        var offsets = new int[segments.Count + 1];
        int cursor = 0;

        for (int i = 0; i < segments.Count; i++)
        {
            offsets[i] = cursor;

            // Capped at the pane, and rounded per segment rather than cumulatively, so that the
            // space reserved on the canvas is exactly the number of pixels the tile can supply.
            // Any mismatch shows up as an unpainted hairline at the seam.
            cursor += Math.Min((int)Math.Round(segments[i].Size * pixelsPerPoint), paneMax);
        }

        offsets[segments.Count] = cursor;

        return offsets;
    }

    /// <summary>Gets the sub-range covered by one tile.</summary>
    private static dynamic GetTileRange(dynamic sheet, dynamic range, CapturePlanner.Segment rowSegment, CapturePlanner.Segment columnSegment)
    {
        dynamic? topLeft = null;
        dynamic? bottomRight = null;

        try
        {
            topLeft = range.Cells[rowSegment.Start, columnSegment.Start];
            bottomRight = range.Cells[rowSegment.Start + rowSegment.Count - 1, columnSegment.Start + columnSegment.Count - 1];

            return sheet.Range[topLeft, bottomRight];
        }
        finally
        {
            ComUtilities.Release(ref bottomRight);
            ComUtilities.Release(ref topLeft);
        }
    }

    /// <summary>
    /// Scrolls the tile's top-left cell to the top-left of the pane. Scrolling rather than selecting
    /// keeps the user's selection intact and avoids drawing a selection border into the capture.
    /// </summary>
    private static void ScrollIntoView(dynamic window, dynamic tileRange)
    {
        try
        {
            window.ScrollRow = Convert.ToInt32(tileRange.Row);
            window.ScrollColumn = Convert.ToInt32(tileRange.Column);
        }
        catch (COMException)
        {
            // Excel clamps scrolling near the sheet edges; the crop follows the actual position.
        }

        Thread.Sleep(ViewSettleMs);
    }

    /// <summary>
    /// Rejects an all-one-color window capture. A real Excel window always has chrome, so a uniform
    /// image means Windows produced nothing - typically a locked desktop or a disconnected Remote
    /// Desktop session.
    /// </summary>
    private static void EnsureWindowRendered(Bitmap shot)
    {
        int stepX = Math.Max(1, shot.Width / 16);
        int stepY = Math.Max(1, shot.Height / 16);
        int first = shot.GetPixel(0, 0).ToArgb();

        for (int y = 0; y < shot.Height; y += stepY)
        {
            for (int x = 0; x < shot.Width; x += stepX)
            {
                if (shot.GetPixel(x, y).ToArgb() != first)
                {
                    return;
                }
            }
        }

        throw new InvalidOperationException(
            "The Excel window rendered as a blank image. " +
            "This happens when the desktop is locked or a Remote Desktop session is disconnected. " +
            "Reconnect to an interactive desktop session and retry the screenshot.");
    }

    private static dynamic? GetActiveSheet(dynamic app)
    {
        try
        {
            return app.ActiveSheet;
        }
        catch (COMException)
        {
            return null;
        }
    }

    private static IntPtr GetExcelWindowHandle(dynamic app)
    {
        IntPtr hwnd = new(Convert.ToInt64(app.Hwnd));

        if (hwnd == IntPtr.Zero)
        {
            throw new InvalidOperationException(
                "Excel did not report a window handle, so its window cannot be captured. " +
                "Make the Excel window visible and retry the screenshot.");
        }

        return hwnd;
    }

    private static void BringExcelToForeground(dynamic app)
    {
        try
        {
            IntPtr hwnd = new(Convert.ToInt64(app.Hwnd));
            if (hwnd != IntPtr.Zero)
            {
                ShowWindow(hwnd, SwRestore);
                SetForegroundWindow(hwnd);
                Thread.Sleep(ViewSettleMs);
            }
        }
        catch (COMException) { }
        catch (InvalidCastException) { }
        catch (FormatException) { }
        catch (OverflowException) { }
    }

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);
}
