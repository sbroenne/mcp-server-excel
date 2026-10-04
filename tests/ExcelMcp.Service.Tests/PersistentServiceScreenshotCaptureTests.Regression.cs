// <copyright file="ScreenshotCommandsTests.Regression.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Regression coverage for screenshot reliability bugs.
/// </summary>
public sealed partial class IsolatedServiceScreenshotTests
{
    [Fact]
    public void CaptureRange_NonActiveOffscreenStyledRange_ProducesVisibleContent()
    {
        var batch = _fixture.BatchToken;
        var sheetName = CreateScreenshotSheetName();

        PopulateHighContrastOffscreenSheet(sheetName, topLeftCell: "AA120");
        _fixture.RegisterSheetForCleanup(sheetName);
        var before = ReadScreenshotState(sheetName, "AA120:AD126");

        var result = _screenshotCommands.CaptureRange(batch, sheetName, "AA120:AD126", ScreenshotQuality.High);

        Assert.Equal(sheetName, result.SheetName);
        Assert.Equal("$AA$120:$AD$126", result.RangeAddress);
        AssertImageLooksPopulated(result);
        Assert.Equal(before, ReadScreenshotState(sheetName, "AA120:AD126"));
    }

    [Fact]
    public void CaptureSheet_NonActiveStyledSheet_ProducesVisibleContent()
    {
        var batch = _fixture.BatchToken;
        var sheetName = CreateScreenshotSheetName();

        PopulateHighContrastOffscreenSheet(sheetName, topLeftCell: "AB90");
        _fixture.RegisterSheetForCleanup(sheetName);
        var before = ReadScreenshotState(sheetName, "AB90:AE96");

        var result = _screenshotCommands.CaptureSheet(batch, sheetName, ScreenshotQuality.High);

        Assert.Equal(sheetName, result.SheetName);
        AssertImageLooksPopulated(result);
        Assert.Equal(before, ReadScreenshotState(sheetName, "AB90:AE96"));
    }

    [Fact]
    public void CaptureRange_RepeatedOffscreenCaptures_AllProduceVisibleContent()
    {
        var batch = _fixture.BatchToken;
        var sheetName = CreateScreenshotSheetName();

        PopulateHighContrastOffscreenSheet(sheetName, topLeftCell: "AA120");
        _fixture.RegisterSheetForCleanup(sheetName);
        var before = ReadScreenshotState(sheetName, "AA120:AD126");

        for (int attempt = 0; attempt < 3; attempt++)
        {
            var result = _screenshotCommands.CaptureRange(batch, sheetName, "AA120:AD126", ScreenshotQuality.High);
            AssertImageLooksPopulated(result);
            Assert.Equal(before, ReadScreenshotState(sheetName, "AA120:AD126"));
        }
    }

    private void PopulateHighContrastOffscreenSheet(string sheetName, string topLeftCell)
    {
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? coverSheet = null;
            Excel.Worksheet? targetSheet = null;
            Excel.Window? window = null;
            Excel.Range? coverCell = null;
            Excel.Interior? coverInterior = null;
            Excel.Range? topLeftRange = null;
            Excel.Range? captureRange = null;
            Excel.Range? totalCell = null;
            Excel.Interior? captureInterior = null;
            Excel.Font? captureFont = null;
            Excel.Font? totalFont = null;

            try
            {
                sheets = ctx.Book.Worksheets;
                coverSheet = (Excel.Worksheet)sheets[1];
                coverSheet.Name = "Cover";
                coverCell = coverSheet.Range["A1"];
                coverCell.Value2 = "Keep this sheet active";
                coverInterior = coverCell.Interior;
                coverInterior.Color = ColorTranslator.ToOle(Color.WhiteSmoke);

                targetSheet = (Excel.Worksheet)sheets.Add(After: coverSheet);
                targetSheet.Name = sheetName;

                topLeftRange = targetSheet.Range[topLeftCell];
                totalCell = topLeftRange.Offset[6, 3];
                captureRange = targetSheet.Range[topLeftRange, totalCell];

                captureInterior = captureRange.Interior;
                captureInterior.Color = ColorTranslator.ToOle(Color.MidnightBlue);
                captureFont = captureRange.Font;
                captureFont.Color = ColorTranslator.ToOle(Color.White);
                captureFont.Bold = true;
                captureRange.RowHeight = 28;
                captureRange.ColumnWidth = 16;

                string[] labels = ["Q1", "Q2", "Q3", "Q4", "FY", "Goal"];
                int[] sales = [120, 165, 140, 190, 615, 650];
                double[] margin = [0.31, 0.42, 0.28, 0.47, 0.37, 0.40];
                string[] status = ["On track", "Ahead", "Watch", "Ahead", "On track", "Stretch"];

                var values = new object[7, 4];
                values[0, 0] = "Quarter";
                values[0, 1] = "Sales";
                values[0, 2] = "Margin";
                values[0, 3] = "Status";
                for (int row = 0; row < labels.Length; row++)
                {
                    values[row + 1, 0] = labels[row];
                    values[row + 1, 1] = (double)sales[row];
                    values[row + 1, 2] = margin[row];
                    values[row + 1, 3] = status[row];
                    Color[] colors = row % 2 == 0
                        ? [Color.SteelBlue, Color.Gold, Color.SeaGreen, Color.IndianRed]
                        : [Color.DarkSlateBlue, Color.Orange, Color.ForestGreen, Color.MediumPurple];
                    for (int column = 0; column < colors.Length; column++)
                    {
                        Excel.Range? cell = null;
                        Excel.Interior? interior = null;
                        try
                        {
                            cell = topLeftRange.Offset[row + 1, column];
                            interior = cell.Interior;
                            interior.Color = ColorTranslator.ToOle(colors[column]);
                        }
                        finally
                        {
                            ComUtilities.Release(ref interior);
                            ComUtilities.Release(ref cell);
                        }
                    }
                }
                captureRange.Value2 = values;
                Assert.Equal(values.Cast<object>(), Assert.IsType<object[,]>((object?)captureRange.Value2).Cast<object>());

                totalFont = totalCell.Font;
                totalFont.Size = 14;

                coverSheet.Activate();
                coverCell.Select();
                window = ctx.App.ActiveWindow;
                window.ScrollRow = 1;
                window.ScrollColumn = 1;
            }
            finally
            {
                ComUtilities.Release(ref totalFont);
                ComUtilities.Release(ref captureFont);
                ComUtilities.Release(ref captureInterior);
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref captureRange);
                ComUtilities.Release(ref totalCell);
                ComUtilities.Release(ref topLeftRange);
                ComUtilities.Release(ref targetSheet);
                ComUtilities.Release(ref coverInterior);
                ComUtilities.Release(ref coverCell);
                ComUtilities.Release(ref coverSheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    private static string CreateScreenshotSheetName() =>
        $"Shot{Guid.NewGuid():N}"[..20];

    private (string Values, string Format, string Cover, string ActiveSheet) ReadScreenshotState(string sheetName, string rangeAddress)
    {
        var values = RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, rangeAddress));
        var cover = RequireSuccess(_commands.GetValues(_fixture.BatchToken, "Cover", "A1"));
        Assert.Equal("Keep this sheet active", cover.Values[0][0]);
        var format = _fixture.Send("rangeformat.get-format", new { sheetName, rangeAddress, view = "stored" });
        Assert.True(format.Success, format.ErrorMessage);
        using var document = JsonDocument.Parse(format.Result!);
        Assert.True(document.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(28, document.RootElement.GetProperty("cellCount").GetInt64());
        var activeSheet = ReadActiveSheetName();
        Assert.Equal("Cover", activeSheet);
        return (JsonSerializer.Serialize(values.Values), document.RootElement.GetProperty("cells").GetRawText(),
            JsonSerializer.Serialize(cover.Values), activeSheet);
    }

    private static void AssertImageLooksPopulated(ScreenshotResult result)
    {
        Assert.True(result.Success, $"Capture failed: {result.ErrorMessage}");
        Assert.Equal("image/png", result.MimeType);
        Assert.NotNull(result.ImageBase64);
        Assert.NotEmpty(result.ImageBase64);

        byte[] imageBytes = Convert.FromBase64String(result.ImageBase64);

        using var stream = new MemoryStream(imageBytes);
        using var bitmap = new Bitmap(stream);
        Assert.Equal(result.Width, bitmap.Width);
        Assert.Equal(result.Height, bitmap.Height);
        AssertImageContainsMarker(result, Color.Gold);
        AssertImageContainsMarker(result, Color.SeaGreen);
        AssertImageContainsMarker(result, Color.IndianRed);

        Assert.True(bitmap.Width >= 200, $"Expected a meaningful screenshot width but got {bitmap.Width}px.");
        Assert.True(bitmap.Height >= 120, $"Expected a meaningful screenshot height but got {bitmap.Height}px.");

        int stepX = Math.Max(1, bitmap.Width / 40);
        int stepY = Math.Max(1, bitmap.Height / 40);
        int sampledPixels = 0;
        int nonWhitePixels = 0;
        int darkPixels = 0;
        HashSet<int> distinctColors = [];

        for (int y = 0; y < bitmap.Height; y += stepY)
        {
            for (int x = 0; x < bitmap.Width; x += stepX)
            {
                Color pixel = bitmap.GetPixel(x, y);
                sampledPixels++;
                distinctColors.Add(pixel.ToArgb());

                if (pixel.R < 245 || pixel.G < 245 || pixel.B < 245)
                {
                    nonWhitePixels++;
                }

                if (pixel.GetBrightness() < 0.85f)
                {
                    darkPixels++;
                }
            }
        }

        Assert.True(nonWhitePixels >= Math.Max(20, sampledPixels / 8), $"Expected visible worksheet content, but only {nonWhitePixels} of {sampledPixels} sampled pixels were non-white.");
        Assert.True(darkPixels >= Math.Max(10, sampledPixels / 12), $"Expected dark formatted content, but only {darkPixels} of {sampledPixels} sampled pixels were dark.");
        Assert.True(distinctColors.Count >= 10, $"Expected multiple visible colors, but sampled only {distinctColors.Count} distinct colors.");
    }
}
