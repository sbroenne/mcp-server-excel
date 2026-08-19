// <copyright file="ScreenshotCommandsTests.WindowCapture.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using System.Drawing;
using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.Screenshot;

/// <summary>
/// Coverage for the window-capture pipeline: tiling of ranges larger than the Excel viewport,
/// restoration of the user's view, and the guarantee that capture leaves the workbook untouched.
/// </summary>
public partial class ScreenshotCommandsTests
{
    private const int Blue = 0xFF0000;
    private const int Red = 0x0000FF;

    [Fact]
    public void CaptureRange_RangeLargerThanViewport_TilesAndIncludesBottomRightContent()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(show: true, operationTimeout: null, testFile);

        PopulateColoredBlock(batch, "Sheet1", "A1:E10", Blue);
        PopulateColoredBlock(batch, "Sheet1", "AV70:AZ80", Red);

        var result = _commands.CaptureRange(batch, sheetName: "Sheet1", rangeAddress: "A1:AZ80", quality: ScreenshotQuality.High);

        Assert.True(result.Success, $"CaptureRange failed: {result.ErrorMessage}");
        Assert.Contains("tiles", result.Message, StringComparison.OrdinalIgnoreCase);

        byte[] imageBytes = Convert.FromBase64String(result.ImageBase64!);
        using var stream = new MemoryStream(imageBytes);
        using var bitmap = new Bitmap(stream);

        Color topLeft = SampleDominantColor(bitmap, 0, 0);
        Color bottomRight = SampleDominantColor(bitmap, bitmap.Width - 1, bitmap.Height - 1);

        Assert.True(
            topLeft.B > topLeft.R,
            $"Expected the blue block at the top-left of the stitched image but found {topLeft}.");
        Assert.True(
            bottomRight.R > bottomRight.B,
            $"Expected the red block at the bottom-right of the stitched image but found {bottomRight}. " +
            "The tiling pass did not reach the end of the range.");
    }

    [Fact]
    public void CaptureRange_RestoresZoomAndScrollPosition()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(show: true, operationTimeout: null, testFile);

        PopulateColoredBlock(batch, "Sheet1", "A100:H120", 65535);
        SetViewState(batch, zoom: 75, scrollRow: 30, scrollColumn: 4);

        var result = _commands.CaptureRange(batch, sheetName: "Sheet1", rangeAddress: "A100:H120");

        Assert.True(result.Success, $"CaptureRange failed: {result.ErrorMessage}");

        (int zoom, int scrollRow, int scrollColumn) = GetViewState(batch);

        Assert.Equal(75, zoom);
        Assert.Equal(30, scrollRow);
        Assert.Equal(4, scrollColumn);
    }

    [Fact]
    public void CaptureRange_LeavesWorkbookUnmodified()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(show: true, operationTimeout: null, testFile);

        PopulateColoredBlock(batch, "Sheet1", "A1:D8", 255);

        var result = _commands.CaptureRange(batch, sheetName: "Sheet1", rangeAddress: "A1:D8");

        Assert.True(result.Success, $"CaptureRange failed: {result.ErrorMessage}");

        (int chartObjects, int shapes, bool cutCopyMode) = GetSheetObjectCounts(batch, "Sheet1");

        Assert.Equal(0, chartObjects);
        Assert.Equal(0, shapes);
        Assert.False(cutCopyMode, "Capture must not leave a pending clipboard operation behind.");
    }

    /// <summary>
    /// Averages a small patch inward from a corner, so grid lines and anti-aliased text do not
    /// dominate the assertion.
    /// </summary>
    private static Color SampleDominantColor(Bitmap bitmap, int cornerX, int cornerY)
    {
        int size = Math.Max(2, Math.Min(bitmap.Width, bitmap.Height) / 40);
        int stepX = cornerX == 0 ? 1 : -1;
        int stepY = cornerY == 0 ? 1 : -1;

        long red = 0;
        long green = 0;
        long blue = 0;
        int samples = 0;

        for (int y = 0; y < size; y++)
        {
            for (int x = 0; x < size; x++)
            {
                Color pixel = bitmap.GetPixel(cornerX + (x * stepX), cornerY + (y * stepY));
                red += pixel.R;
                green += pixel.G;
                blue += pixel.B;
                samples++;
            }
        }

        return Color.FromArgb((int)(red / samples), (int)(green / samples), (int)(blue / samples));
    }

    private static void SetViewState(IExcelBatch batch, int zoom, int scrollRow, int scrollColumn)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? window = null;
            try
            {
                window = ctx.App.ActiveWindow;
                window.Zoom = zoom;
                window.ScrollRow = scrollRow;
                window.ScrollColumn = scrollColumn;
            }
            finally
            {
                ComUtilities.Release(ref window);
            }
        });
    }

    private static (int Zoom, int ScrollRow, int ScrollColumn) GetViewState(IExcelBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? window = null;
            try
            {
                window = ctx.App.ActiveWindow;
                int zoom = Convert.ToInt32((object)window.Zoom, CultureInfo.InvariantCulture);
                int scrollRow = Convert.ToInt32((object)window.ScrollRow, CultureInfo.InvariantCulture);
                int scrollColumn = Convert.ToInt32((object)window.ScrollColumn, CultureInfo.InvariantCulture);

                return (zoom, scrollRow, scrollColumn);
            }
            finally
            {
                ComUtilities.Release(ref window);
            }
        });
    }

    private static (int ChartObjects, int Shapes, bool CutCopyMode) GetSheetObjectCounts(IExcelBatch batch, string sheetName)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? charts = null;
            dynamic? shapes = null;
            try
            {
                sheet = ctx.Book.Worksheets[sheetName];
                charts = sheet.ChartObjects();
                shapes = sheet.Shapes;

                int chartCount = Convert.ToInt32((object)charts.Count, CultureInfo.InvariantCulture);
                int shapeCount = Convert.ToInt32((object)shapes.Count, CultureInfo.InvariantCulture);
                bool cutCopyMode = Convert.ToInt32((object)ctx.App.CutCopyMode, CultureInfo.InvariantCulture) != 0;

                return (chartCount, shapeCount, cutCopyMode);
            }
            finally
            {
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref charts);
                ComUtilities.Release(ref sheet);
            }
        });
    }
}
