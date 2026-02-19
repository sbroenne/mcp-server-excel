using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Core.Commands.Screenshot;

/// <summary>
/// Implementation of screenshot commands using Excel COM CopyPicture + ChartObject.Export.
/// </summary>
public class ScreenshotCommands : IScreenshotCommands
{
    // Excel COM constants
    private const int XlScreen = 1;      // xlScreen - required for CopyPicture to render correctly
    private const int XlBitmap = 2;      // xlBitmap

    // CopyPicture retry configuration
    // After Save or large operations, Excel rendering can take several seconds.
    // Retries with exponential backoff: 500ms, 1000ms, 1500ms, 2000ms, 2500ms, 3000ms, 3500ms
    private const int CopyPictureMaxRetries = 10;
    private const int CopyPictureRetryDelayMs = 700;

    /// <summary>
    /// Captures a specific range as a PNG image.
    /// </summary>
    public ScreenshotResult CaptureRange(IExcelBatch batch, string? sheetName = null, string rangeAddress = "A1:Z30")
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

                return ExportRangeAsImage(ctx.App, sheet, range, actualSheet, actualRange);
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <summary>
    /// Captures the entire used area of a worksheet as a PNG image.
    /// If UsedRange exceeds 500 rows or 50 columns, it is capped to avoid
    /// CopyPicture failures on sheets with formatting extending far beyond data.
    /// </summary>
    public ScreenshotResult CaptureSheet(IExcelBatch batch, string? sheetName = null)
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

                int rows = (int)usedRange.Rows.Count;
                int cols = (int)usedRange.Columns.Count;

                const int maxRows = 500;
                const int maxCols = 50;

                if (rows > maxRows || cols > maxCols)
                {
                    // Cap the range to avoid CopyPicture failures on enormous ranges
                    int startRow = (int)usedRange.Row;
                    int startCol = (int)usedRange.Column;
                    int endRow = startRow + Math.Min(rows, maxRows) - 1;
                    int endCol = startCol + Math.Min(cols, maxCols) - 1;
                    captureRange = sheet.Range[sheet.Cells[startRow, startCol], sheet.Cells[endRow, endCol]];
                }

                dynamic rangeToCapture = captureRange ?? usedRange;
                string actualRange = rangeToCapture.Address?.ToString() ?? "A1";

                return ExportRangeAsImage(ctx.App, sheet, rangeToCapture, actualSheet, actualRange);
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
    /// Exports a range as a PNG image using CopyPicture + ChartObject.Export.
    /// CopyPicture requires Excel to be visible for rendering. If Excel is hidden,
    /// we temporarily show it, capture, then restore the previous visibility state.
    ///
    /// CRITICAL: After Save/large operations, Excel needs rendering time even if already visible.
    /// This method includes delays to ensure Excel is fully rendered before capture.
    /// </summary>
    private static ScreenshotResult ExportRangeAsImage(dynamic app, dynamic sheet, dynamic range, string sheetName, string rangeAddress)
    {
        dynamic? chartObjects = null;
        dynamic? chartObject = null;
        dynamic? chart = null;
        string? tempFile = null;
        bool wasVisible = false;

        try
        {
            // CopyPicture requires Excel to be visible for UI rendering
            wasVisible = (bool)app.Visible;
            if (!wasVisible)
            {
                app.Visible = true;

                // Excel needs time to initialize its rendering pipeline after
                // becoming visible. Without this, CopyPicture fails with
                // "Unable to get the CopyPicture property" or crashes the process
                // with RPC_S_SERVER_UNAVAILABLE (0x800706BA) under rapid cycling.
                Thread.Sleep(2000);
            }
            else
            {
                // Excel is already visible, but may still be rendering from previous operations.
                // Slightly longer delay to allow rendering pipeline to fully settle.
                Thread.Sleep(1000);
            }

            // Try to activate the Excel window to ensure it has focus for proper rendering.
            // This helps ensure content is fully rendered before capture.
            try
            {
                app.Activate();
                // Allow activation and message pump to settle
                Thread.Sleep(500);
            }
            catch
            {
                // Activate may fail in some contexts (minimized, headless, etc.) - ignore
            }

            // Get range dimensions for the chart
            double width = Convert.ToDouble(range.Width);
            double height = Convert.ToDouble(range.Height);

            // Cap dimensions to avoid huge images (max ~4096px equivalent at 96 DPI)
            // Excel Width/Height are in points (1 point = 1.333 pixels at 96 DPI)
            const double maxPoints = 3072; // ~4096px
            if (width > maxPoints || height > maxPoints)
            {
                double scale = Math.Min(maxPoints / width, maxPoints / height);
                width *= scale;
                height *= scale;
            }

            // Copy range as picture (with retry — CopyPicture is clipboard-dependent
            // and intermittently fails when Excel is still rendering after chart/table operations)
            CopyPictureWithRetry(range);

            // Create a temporary ChartObject to paste into and export
            chartObjects = sheet.ChartObjects();
            chartObject = chartObjects.Add(0, 0, width, height);
            chart = chartObject.Chart;

            // Paste the copied picture into the chart
            chart.Paste();

            // Clear clipboard immediately after paste — releases clipboard for subsequent screenshot calls
            // (otherwise marching ants remain and next CopyPicture may fail with clipboard contention)
            try { app.CutCopyMode = false; } catch { /* best effort */ }

            // Export to temp PNG file
            tempFile = Path.Combine(Path.GetTempPath(), $"excelmcp-screenshot-{Guid.NewGuid():N}.png");
            chart.Export(tempFile, "PNG");

            // Read and convert to base64
            byte[] imageBytes = File.ReadAllBytes(tempFile);
            string base64 = Convert.ToBase64String(imageBytes);

            // Get actual pixel dimensions from the PNG header
            (int pixelWidth, int pixelHeight) = GetPngDimensions(imageBytes);

            return new ScreenshotResult
            {
                Success = true,
                ImageBase64 = base64,
                MimeType = "image/png",
                Width = pixelWidth,
                Height = pixelHeight,
                SheetName = sheetName,
                RangeAddress = rangeAddress,
                Message = $"Captured {rangeAddress} on '{sheetName}' ({pixelWidth}x{pixelHeight}px)"
            };
        }
        finally
        {
            // Clean up temp ChartObject from the worksheet
            if (chartObject != null)
            {
                try { chartObject.Delete(); } catch { /* best effort */ }
            }

            ComUtilities.Release(ref chart);
            ComUtilities.Release(ref chartObject);
            ComUtilities.Release(ref chartObjects);

            // Restore Excel visibility if we changed it
            if (!wasVisible)
            {
                try { app.Visible = false; } catch { /* best effort */ }
            }

            // Clean up temp file
            if (tempFile != null && File.Exists(tempFile))
            {
                try { File.Delete(tempFile); } catch { /* best effort */ }
            }
        }
    }

    /// <summary>
    /// Calls CopyPicture with retry logic. CopyPicture uses the clipboard and
    /// intermittently fails with COMException when Excel is busy rendering
    /// (e.g., after chart/table operations). Retries with increasing delay.
    /// </summary>
    private static void CopyPictureWithRetry(dynamic range)
    {
        for (int attempt = 0; attempt < CopyPictureMaxRetries; attempt++)
        {
            try
            {
                range.CopyPicture(XlScreen, XlBitmap);
                return;
            }
            catch (COMException) when (attempt < CopyPictureMaxRetries - 1)
            {
                Thread.Sleep(CopyPictureRetryDelayMs * (attempt + 1));
            }
        }
    }

    /// <summary>
    /// Reads width and height from PNG file header (IHDR chunk).
    /// PNG format: 8-byte signature, then IHDR chunk with width (4 bytes) and height (4 bytes).
    /// </summary>
    private static (int width, int height) GetPngDimensions(byte[] data)
    {
        if (data.Length < 24)
            return (0, 0);

        // PNG IHDR starts at byte 16 (after 8-byte signature + 4-byte length + 4-byte "IHDR")
        int width = (data[16] << 24) | (data[17] << 16) | (data[18] << 8) | data[19];
        int height = (data[20] << 24) | (data[21] << 16) | (data[22] << 8) | data[23];

        return (width, height);
    }
}
