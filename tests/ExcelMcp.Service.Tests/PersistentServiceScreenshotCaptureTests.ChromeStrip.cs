// <copyright file="ScreenshotCommandsTests.ChromeStrip.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Regression coverage for issue #777 follow-up: captures must contain only worksheet grid, never
/// the Excel window chrome below it (horizontal scroll bar, sheet tabs, status bar).
///
/// <c>Window.UsableHeight</c> measures the whole workspace including that chrome, so sizing a tile
/// from it reaches past the bottom of the grid. Each test fills the captured range with a solid
/// colour, which makes any leaked chrome trivially detectable: every pixel row of a correct capture
/// is that colour, so a row that is not is chrome.
/// </summary>
public sealed partial class IsolatedServiceScreenshotTests
{
    /// <summary>Solid black fill; chrome is far lighter than this at every Office theme.</summary>
    private const int SolidBlack = 0x000000;

    /// <summary>Row brightness above which a row cannot be part of a solid black fill.</summary>
    private const int ChromeBrightnessThreshold = 100;

    /// <summary>
    /// Brightness at or above which a row is the untouched white canvas rather than Excel chrome.
    /// Excel's chrome (scroll bar, tab strip) is grey and measures well below this.
    /// </summary>
    private const int UnpaintedCanvasBrightness = 250;

    /// <summary>
    /// Pixel rows at the very bottom edge that may legitimately stay unpainted: Excel rounds each
    /// rendered row to whole pixels, so a tall range renders a shade shorter than the exact
    /// point-to-pixel arithmetic predicts.
    /// </summary>
    private const int EdgeRoundingTolerance = 3;

    [Fact]
    public void CaptureRange_TallRangeSpanningMultipleTiles_ContainsNoWindowChrome()
    {
        var batch = _fixture.BatchToken;

        FillSolid("Sheet1", "A1:A200", SolidBlack);

        var result = _screenshotCommands.CaptureRange(batch, sheetName: "Sheet1", rangeAddress: "A1:A200", quality: ScreenshotQuality.High);

        Assert.True(result.Success, $"CaptureRange failed: {result.ErrorMessage}");

        AssertNoChromeBands(result.ImageBase64!, "A1:A200");
    }

    [Fact]
    public void CaptureRange_RangeFillingTheViewport_ContainsNoWindowChrome()
    {
        var batch = _fixture.BatchToken;

        FillSolid("Sheet1", "A1:A45", SolidBlack);

        var result = _screenshotCommands.CaptureRange(batch, sheetName: "Sheet1", rangeAddress: "A1:A45", quality: ScreenshotQuality.High);

        Assert.True(result.Success, $"CaptureRange failed: {result.ErrorMessage}");

        AssertNoChromeBands(result.ImageBase64!, "A1:A45");
    }

    [Theory]
    [InlineData(590)]
    [InlineData(610)]
    [InlineData(630)]
    public void CaptureRange_NonIntegralRowPixelHeight_ContainsNoWindowChrome(int windowHeight)
    {
        var batch = _fixture.BatchToken;
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            ctx.App.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlNormal;
            ctx.App.Width = 900;
            ctx.App.Height = windowHeight;
        });
        FillSolid("Sheet1", "A1:A45", SolidBlack);

        var result = _screenshotCommands.CaptureRange(batch, "Sheet1", "A1:A45", ScreenshotQuality.High);

        Assert.True(result.Success, result.ErrorMessage);
        AssertNoChromeBands(result.ImageBase64!, $"A1:A45, window height {windowHeight}");
    }

    [Theory]
    [InlineData(45, 3)]
    [InlineData(200, 12)]
    public void CaptureRange_VariedCellSizes_PreservesEveryRowAndColumn(int rowCount, int columnCount)
    {
        var batch = _fixture.BatchToken;
        string address = $"A1:{(char)('A' + columnCount - 1)}{rowCount}";
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            ctx.App.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlNormal;
            ctx.App.Width = 900;
            ctx.App.Height = 630;
            dynamic? sheet = null;
            dynamic? cells = null;
            dynamic? parkingCell = null;
            try
            {
                sheet = ctx.Book.Worksheets["Sheet1"];
                sheet.Activate();
                cells = sheet.Cells;
                for (int row = 1; row <= rowCount; row++)
                {
                    for (int column = 1; column <= columnCount; column++)
                    {
                        dynamic? cell = null;
                        dynamic? interior = null;
                        try
                        {
                            cell = cells[row, column];
                            if (column == 1) cell.RowHeight = 11.25 + row % 3 * 3.75;
                            if (row == 1) cell.ColumnWidth = 7.5 + column % 3 * 2.25;
                            interior = cell.Interior;
                            interior.Color = ColorTranslator.ToOle((row + column) % 2 == 0 ? Color.Navy : Color.Maroon);
                        }
                        finally
                        {
                            ComUtilities.Release(ref interior);
                            ComUtilities.Release(ref cell);
                        }
                    }
                }
                parkingCell = cells[1, columnCount + 2];
                parkingCell.Select();
            }
            finally
            {
                ComUtilities.Release(ref parkingCell);
                ComUtilities.Release(ref cells);
                ComUtilities.Release(ref sheet);
            }
        });

        var result = _screenshotCommands.CaptureRange(batch, "Sheet1", address, ScreenshotQuality.High);

        Assert.True(result.Success, result.ErrorMessage);
        AssertNoChromeBands(result.ImageBase64!, address);
        using var stream = new MemoryStream(Convert.FromBase64String(result.ImageBase64!));
        using var bitmap = new Bitmap(stream);
        Assert.Equal(rowCount, CountColorBands(bitmap.Height, y => bitmap.GetPixel(bitmap.Width / 2, y)));
        Assert.Equal(columnCount, CountColorBands(bitmap.Width, x => bitmap.GetPixel(x, bitmap.Height / 2)));
    }

    private static int CountColorBands(int length, Func<int, Color> getPixel)
    {
        int bands = 0;
        int previous = -1;
        for (int i = 0; i < length; i++)
        {
            Color color = getPixel(i);
            int current = color.B > 100 && color.R < 20 ? 0 : color.R > 100 && color.B < 20 ? 1 : -1;
            if (current >= 0 && current != previous)
            {
                bands++;
                previous = current;
            }
        }
        return bands;
    }

    /// <summary>
    /// Fails when any pixel row of the capture is lighter than a solid black fill, which means the
    /// window chrome below the grid was captured instead of worksheet content.
    /// </summary>
    private static void AssertNoChromeBands(string base64, string rangeAddress)
    {
        byte[] imageBytes = Convert.FromBase64String(base64);
        using var stream = new MemoryStream(imageBytes);
        using var bitmap = new Bitmap(stream);

        var lightRows = new List<string>();

        for (int y = 0; y < bitmap.Height; y++)
        {
            long total = 0;
            int samples = 0;

            for (int x = 0; x < bitmap.Width; x++)
            {
                Color pixel = bitmap.GetPixel(x, y);
                total += (pixel.R + pixel.G + pixel.B) / 3;
                samples++;
            }

            int average = (int)(total / samples);

            if (average <= ChromeBrightnessThreshold)
            {
                continue;
            }

            bool unpaintedBottomEdge =
                y >= bitmap.Height - EdgeRoundingTolerance && average >= UnpaintedCanvasBrightness;

            if (!unpaintedBottomEdge)
            {
                lightRows.Add(string.Create(CultureInfo.InvariantCulture, $"y={y} avg={average}"));
            }
        }

        Assert.True(
            lightRows.Count == 0,
            string.Create(
                CultureInfo.InvariantCulture,
                $"Capture of {rangeAddress} ({bitmap.Width}x{bitmap.Height}) contains {lightRows.Count} pixel row(s) that are not the solid fill, " +
                $"which means Excel window chrome was captured below the grid: {string.Join(", ", lightRows.Take(12))}"));
    }

    private void FillSolid(string sheetName, string rangeAddress, int fillColor)
    {
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? parkingCell = null;
            try
            {
                sheet = ctx.Book.Worksheets[sheetName];
                range = sheet.Range[rangeAddress];
                range.Interior.Color = fillColor;

                // A real screenshot includes the active-cell outline, which would otherwise be drawn
                // over the fill and read as a light row. Park the selection outside the capture.
                sheet.Activate();
                parkingCell = sheet.Range["E1"];
                parkingCell.Select();
            }
            finally
            {
                ComUtilities.Release(ref parkingCell);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
    }
}
