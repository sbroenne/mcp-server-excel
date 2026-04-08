// <copyright file="ScreenshotCommandsTests.Regression.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using System.Drawing;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.Screenshot;

/// <summary>
/// Regression coverage for screenshot reliability bugs.
/// </summary>
public partial class ScreenshotCommandsTests
{
    [Fact]
    public void CaptureRange_NonActiveOffscreenStyledRange_ProducesVisibleContent()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(show: true, operationTimeout: null, testFile);

        PopulateHighContrastOffscreenSheet(batch, sheetName: "Report", topLeftCell: "AA120");

        var result = _commands.CaptureRange(batch, sheetName: "Report", rangeAddress: "AA120:AD126", quality: ScreenshotQuality.High);

        Assert.Equal("Report", result.SheetName);
        Assert.Equal("$AA$120:$AD$126", result.RangeAddress);
        AssertImageLooksPopulated(result);
    }

    [Fact]
    public void CaptureSheet_NonActiveStyledSheet_ProducesVisibleContent()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(show: true, operationTimeout: null, testFile);

        PopulateHighContrastOffscreenSheet(batch, sheetName: "Dashboard", topLeftCell: "AB90");

        var result = _commands.CaptureSheet(batch, sheetName: "Dashboard", quality: ScreenshotQuality.High);

        Assert.Equal("Dashboard", result.SheetName);
        AssertImageLooksPopulated(result);
    }

    [Fact]
    public void CaptureRange_RepeatedOffscreenCaptures_AllProduceVisibleContent()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(show: true, operationTimeout: null, testFile);

        PopulateHighContrastOffscreenSheet(batch, sheetName: "Report", topLeftCell: "AA120");

        for (int attempt = 0; attempt < 3; attempt++)
        {
            var result = _commands.CaptureRange(batch, sheetName: "Report", rangeAddress: "AA120:AD126", quality: ScreenshotQuality.High);
            AssertImageLooksPopulated(result);
        }
    }

    private static void PopulateHighContrastOffscreenSheet(IExcelBatch batch, string sheetName, string topLeftCell)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? coverSheet = null;
            dynamic? targetSheet = null;
            dynamic? window = null;
            dynamic? topLeftRange = null;
            dynamic? captureRange = null;
            dynamic? titleCell = null;
            dynamic? totalCell = null;

            try
            {
                coverSheet = ctx.Book.Worksheets[1];
                coverSheet.Name = "Cover";
                coverSheet.Range["A1"].Value2 = "Keep this sheet active";
                coverSheet.Range["A1"].Interior.Color = ColorTranslator.ToOle(Color.WhiteSmoke);

                targetSheet = ctx.Book.Worksheets.Add(After: coverSheet);
                targetSheet.Name = sheetName;

                topLeftRange = targetSheet.Range[topLeftCell];
                captureRange = targetSheet.Range[topLeftRange, topLeftRange.Offset[6, 3]];

                captureRange.Interior.Color = ColorTranslator.ToOle(Color.MidnightBlue);
                captureRange.Font.Color = ColorTranslator.ToOle(Color.White);
                captureRange.Font.Bold = true;
                captureRange.RowHeight = 28;
                captureRange.ColumnWidth = 16;

                titleCell = topLeftRange;
                totalCell = topLeftRange.Offset[6, 3];

                titleCell.Value2 = "Quarter";
                titleCell.Offset[0, 1].Value2 = "Sales";
                titleCell.Offset[0, 2].Value2 = "Margin";
                titleCell.Offset[0, 3].Value2 = "Status";

                string[] labels = ["Q1", "Q2", "Q3", "Q4", "FY", "Goal"];
                int[] sales = [120, 165, 140, 190, 615, 650];
                double[] margin = [0.31, 0.42, 0.28, 0.47, 0.37, 0.40];
                string[] status = ["On track", "Ahead", "Watch", "Ahead", "On track", "Stretch"];

                for (int row = 0; row < labels.Length; row++)
                {
                    dynamic? labelCell = null;
                    dynamic? salesCell = null;
                    dynamic? marginCell = null;
                    dynamic? statusCell = null;

                    try
                    {
                        labelCell = topLeftRange.Offset[row + 1, 0];
                        salesCell = topLeftRange.Offset[row + 1, 1];
                        marginCell = topLeftRange.Offset[row + 1, 2];
                        statusCell = topLeftRange.Offset[row + 1, 3];

                        labelCell.Value2 = labels[row];
                        salesCell.Value2 = sales[row];
                        marginCell.Value2 = margin[row];
                        statusCell.Value2 = status[row];

                        labelCell.Interior.Color = ColorTranslator.ToOle(row % 2 == 0 ? Color.SteelBlue : Color.DarkSlateBlue);
                        salesCell.Interior.Color = ColorTranslator.ToOle(row % 2 == 0 ? Color.Gold : Color.Orange);
                        marginCell.Interior.Color = ColorTranslator.ToOle(row % 2 == 0 ? Color.SeaGreen : Color.ForestGreen);
                        statusCell.Interior.Color = ColorTranslator.ToOle(row % 2 == 0 ? Color.IndianRed : Color.MediumPurple);
                    }
                    finally
                    {
                        ComUtilities.Release(ref statusCell);
                        ComUtilities.Release(ref marginCell);
                        ComUtilities.Release(ref salesCell);
                        ComUtilities.Release(ref labelCell);
                    }
                }

                totalCell.Font.Size = 14;

                coverSheet.Activate();
                coverSheet.Range["A1"].Select();
                window = ctx.App.ActiveWindow;
                if (window != null)
                {
                    window.ScrollRow = 1;
                    window.ScrollColumn = 1;
                }
            }
            finally
            {
                ComUtilities.Release(ref totalCell);
                ComUtilities.Release(ref titleCell);
                ComUtilities.Release(ref captureRange);
                ComUtilities.Release(ref topLeftRange);
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref targetSheet);
                ComUtilities.Release(ref coverSheet);
            }
        });
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
