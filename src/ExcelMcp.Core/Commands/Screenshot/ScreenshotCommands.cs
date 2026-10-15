using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

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
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            try
            {
                sheet = ResolveWorksheet(ctx.Book, sheetName);

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
    /// Captures the entire used area of a worksheet and its embedded charts as an image.
    /// Large sheets are tiled and may be truncated by the capture tile limit, which is reported in
    /// the result message.
    /// </summary>
    public ScreenshotResult CaptureSheet(IExcelBatch batch, string? sheetName = null, ScreenshotQuality quality = ScreenshotQuality.Medium)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? usedRange = null;
            Excel.Range? captureRange = null;
            try
            {
                sheet = ResolveWorksheet(ctx.Book, sheetName);

                usedRange = sheet.UsedRange;
                string actualSheet = sheet.Name?.ToString() ?? "Sheet1";

                captureRange = GetCaptureRangeIncludingCharts(sheet, usedRange, ct);

                string actualRange = captureRange.Address?.ToString() ?? "A1";

                return ExportRangeAsImage(ctx.App, sheet, captureRange, actualSheet, actualRange, quality);
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
    /// Gets the smallest cell range containing the used cells and all embedded chart shapes.
    /// Chart objects do not expand Excel's UsedRange, but their anchor cells describe the area that
    /// must be visible for a window screenshot to include them.
    /// </summary>
    private static Excel.Range GetCaptureRangeIncludingCharts(Excel.Worksheet sheet, Excel.Range usedRange, CancellationToken ct)
    {
        Excel.Range? usedRows = null;
        Excel.Range? usedColumns = null;
        Excel.Shapes? shapes = null;
        Excel.Shape? shape = null;
        Excel.Range? topLeftCell = null;
        Excel.Range? bottomRightCell = null;

        try
        {
            int firstRow = Convert.ToInt32(usedRange.Row);
            int firstColumn = Convert.ToInt32(usedRange.Column);
            usedRows = usedRange.Rows;
            usedColumns = usedRange.Columns;
            int lastRow = firstRow + Convert.ToInt32(usedRows.Count) - 1;
            int lastColumn = firstColumn + Convert.ToInt32(usedColumns.Count) - 1;

            shapes = sheet.Shapes;
            int shapeCount = Convert.ToInt32(shapes.Count);

            for (int index = 1; index <= shapeCount; index++)
            {
                ct.ThrowIfCancellationRequested();

                try
                {
                    shape = shapes.Item(index);

                    // msoChart
                    if (GetShapeType(shape) != 3)
                    {
                        continue;
                    }

                    // msoFalse: hidden charts are not rendered, so they must not expand the capture area
                    if (!IsShapeVisible(shape))
                    {
                        continue;
                    }

                    topLeftCell = shape.TopLeftCell;
                    bottomRightCell = shape.BottomRightCell;

                    firstRow = Math.Min(firstRow, Convert.ToInt32(topLeftCell.Row));
                    firstColumn = Math.Min(firstColumn, Convert.ToInt32(topLeftCell.Column));
                    lastRow = Math.Max(lastRow, Convert.ToInt32(bottomRightCell.Row));
                    lastColumn = Math.Max(lastColumn, Convert.ToInt32(bottomRightCell.Column));
                }
                finally
                {
                    ComUtilities.Release(ref bottomRightCell);
                    ComUtilities.Release(ref topLeftCell);
                    ComUtilities.Release(ref shape);
                }
            }

            return GetRange(sheet, firstRow, firstColumn, lastRow, lastColumn);
        }
        finally
        {
            ComUtilities.Release(ref shapes);
            ComUtilities.Release(ref usedColumns);
            ComUtilities.Release(ref usedRows);
        }
    }

    /// <summary>Gets a worksheet range by its inclusive cell boundaries.</summary>
    private static Excel.Range GetRange(Excel.Worksheet sheet, int firstRow, int firstColumn, int lastRow, int lastColumn)
    {
        Excel.Range? cells = null;
        Excel.Range? topLeft = null;
        Excel.Range? bottomRight = null;

        try
        {
            cells = sheet.Cells;
            topLeft = (Excel.Range)cells[firstRow, firstColumn];
            bottomRight = (Excel.Range)cells[lastRow, lastColumn];

            return sheet.Range[topLeft, bottomRight];
        }
        finally
        {
            ComUtilities.Release(ref bottomRight);
            ComUtilities.Release(ref topLeft);
            ComUtilities.Release(ref cells);
        }
    }

    /// <summary>
    /// Captures a range by photographing the Excel window and cropping to the range rectangle.
    ///
    /// The window must be visible for Windows to render it, so a hidden or minimized Excel is
    /// temporarily shown. Zoom and scroll position are adjusted to bring the range into view and
    /// are restored afterwards.
    /// </summary>
    private static ScreenshotResult ExportRangeAsImage(Excel.Application app, Excel.Worksheet sheet, Excel.Range range, string sheetName, string rangeAddress, ScreenshotQuality quality)
    {
        return DpiAwareness.Execute<ScreenshotResult>(() => CaptureRangeImage(app, sheet, range, sheetName, rangeAddress, quality));
    }

    private static ScreenshotResult CaptureRangeImage(Excel.Application app, Excel.Worksheet sheet, Excel.Range range, string sheetName, string rangeAddress, ScreenshotQuality quality)
    {
        Excel.Window? window = null;
        object? previousSheet = null;
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

            originalWindowState = Convert.ToInt32(app.WindowState, System.Globalization.CultureInfo.InvariantCulture);
            if (originalWindowState == XlMinimized)
            {
                app.WindowState = (Excel.XlWindowState)XlNormal;
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
                try
                {
                    switch (previousSheet)
                    {
                        case Excel.Worksheet worksheet: worksheet.Activate(); break;
                        case Excel.Chart chart: chart.Activate(); break;
                        case Excel.DialogSheet dialog: dialog.Activate(); break;
                    }
                }
                catch (COMException) { }
            }

            if (windowStateChanged)
            {
                try { app.WindowState = (Excel.XlWindowState)originalWindowState; } catch (COMException) { }
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
    private static Bitmap CaptureTiles(Excel.Application app, Excel.Window window, Excel.Range range, ref CapturePlanner.CapturePlan plan)
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

        int[] columnOffsets = BuildPixelOffsets(range, false, plan.ColumnSegments, pixelsPerPoint, paneMaxWidth);
        int[] rowOffsets = BuildPixelOffsets(range, true, plan.RowSegments, pixelsPerPoint, paneMaxHeight);

        int totalWidth = Math.Max(1, columnOffsets[^1]);
        int totalHeight = Math.Max(1, rowOffsets[^1]);

        var canvas = new Bitmap(totalWidth, totalHeight, PixelFormat.Format32bppArgb);
        Excel.Worksheet? sheet = range.Worksheet;

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

                    Excel.Range? tileRange = null;
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
                        int offsetX = Convert.ToInt32(Math.Round((Convert.ToDouble((object)tileRange.Left, System.Globalization.CultureInfo.InvariantCulture) - GetColumnLeft(sheet, window.ScrollColumn)) * pixelsPerPoint), System.Globalization.CultureInfo.InvariantCulture);
                        int offsetY = Convert.ToInt32(Math.Round((Convert.ToDouble((object)tileRange.Top, System.Globalization.CultureInfo.InvariantCulture) - GetRowTop(sheet, window.ScrollRow)) * pixelsPerPoint), System.Globalization.CultureInfo.InvariantCulture);

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
    private static PaneOrigin MeasurePaneOrigin(Excel.Window window)
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
    private static double GetRowTop(Excel.Worksheet sheet, int row)
    {
        Excel.Range? cells = null;
        Excel.Range? cell = null;
        try
        {
            cells = sheet.Cells;
            cell = (Excel.Range)cells[row, 1];
            return Convert.ToDouble(cell.Top);
        }
        finally
        {
            ComUtilities.Release(ref cell);
            ComUtilities.Release(ref cells);
        }
    }

    /// <summary>Gets the worksheet point offset of a column's left edge.</summary>
    private static double GetColumnLeft(Excel.Worksheet sheet, int column)
    {
        Excel.Range? cells = null;
        Excel.Range? cell = null;
        try
        {
            cells = sheet.Cells;
            cell = (Excel.Range)cells[1, column];
            return Convert.ToDouble(cell.Left);
        }
        finally
        {
            ComUtilities.Release(ref cell);
            ComUtilities.Release(ref cells);
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
    private static int[] BuildPixelOffsets(Excel.Range range, bool rows, IReadOnlyList<CapturePlanner.Segment> segments, double pixelsPerPoint, int paneMax)
    {
        var offsets = new int[segments.Count + 1];
        int cursor = 0;
        Excel.Range? items = null;
        try
        {
            items = rows ? range.Rows : range.Columns;
            for (int i = 0; i < segments.Count; i++)
            {
                offsets[i] = cursor;
                int renderedSize = 0;
                var segment = segments[i];
                for (int index = segment.Start; index < segment.Start + segment.Count; index++)
                {
                    Excel.Range? item = null;
                    try
                    {
                        item = (Excel.Range)items[index];
                        double extent = rows ? Convert.ToDouble(item.Height) : Convert.ToDouble(item.Width);
                        // Excel rounds each rendered cell to whole pixels, not the summed range height.
                        renderedSize += (int)Math.Round(extent * pixelsPerPoint, MidpointRounding.AwayFromZero);
                    }
                    finally
                    {
                        ComUtilities.Release(ref item);
                    }
                }

                cursor += Math.Min(renderedSize, paneMax);
            }
        }
        finally
        {
            ComUtilities.Release(ref items);
        }

        offsets[segments.Count] = cursor;

        return offsets;
    }

    /// <summary>Gets the sub-range covered by one tile.</summary>
    private static Excel.Range GetTileRange(Excel.Worksheet sheet, Excel.Range range, CapturePlanner.Segment rowSegment, CapturePlanner.Segment columnSegment)
    {
        Excel.Range? cells = null;
        Excel.Range? topLeft = null;
        Excel.Range? bottomRight = null;

        try
        {
            cells = range.Cells;
            topLeft = (Excel.Range)cells[rowSegment.Start, columnSegment.Start];
            bottomRight = (Excel.Range)cells[rowSegment.Start + rowSegment.Count - 1, columnSegment.Start + columnSegment.Count - 1];

            return sheet.Range[topLeft, bottomRight];
        }
        finally
        {
            ComUtilities.Release(ref bottomRight);
            ComUtilities.Release(ref topLeft);
            ComUtilities.Release(ref cells);
        }
    }

    /// <summary>
    /// Scrolls the tile's top-left cell to the top-left of the pane. Scrolling rather than selecting
    /// keeps the user's selection intact and avoids drawing a selection border into the capture.
    /// </summary>
    private static void ScrollIntoView(Excel.Window window, Excel.Range tileRange)
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

    private static object? GetActiveSheet(Excel.Application app)
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

    private static IntPtr GetExcelWindowHandle(Excel.Application app)
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

    private static void BringExcelToForeground(Excel.Application app)
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

    private static int GetShapeType(Excel.Shape shape)
    {
        // PIA gap: Shape.Type returns Office.Core.MsoShapeType, requiring the unavailable Office-core assembly (CS0012).
        return Convert.ToInt32(((dynamic)shape).Type, System.Globalization.CultureInfo.InvariantCulture);
    }

    private static bool IsShapeVisible(Excel.Shape shape)
    {
        // PIA gap: Shape.Visible returns Office.Core.MsoTriState, requiring the unavailable Office-core assembly (CS0012).
        return Convert.ToInt32(((dynamic)shape).Visible, System.Globalization.CultureInfo.InvariantCulture) != 0;
    }

    private static Excel.Worksheet ResolveWorksheet(Excel.Workbook book, string? sheetName)
    {
        Excel.Sheets? sheets = null;
        object? selected = null;
        try
        {
            if (string.IsNullOrWhiteSpace(sheetName))
                selected = book.ActiveSheet;
            else
            {
                sheets = book.Worksheets;
                selected = sheets[sheetName];
            }

            if (selected is not Excel.Worksheet worksheet)
                throw new InvalidOperationException("The selected sheet is not a worksheet and has no cell range to capture.");
            selected = null;
            return worksheet;
        }
        finally
        {
            ComUtilities.Release(ref selected);
            ComUtilities.Release(ref sheets);
        }
    }

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);
}
