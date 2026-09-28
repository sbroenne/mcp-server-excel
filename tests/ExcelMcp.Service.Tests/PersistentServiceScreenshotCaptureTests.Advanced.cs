// <copyright file="ScreenshotCommandsTests.Capture.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for CaptureRange and CaptureSheet operations.
/// These exercise the window-capture pipeline end to end.
/// </summary>
[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Screenshot")]
public sealed partial class IsolatedServiceScreenshotTests :
    IsolatedServiceScreenshotTestBase
{
    /// <summary>
    /// Helper: populates a test file with sample data and optionally a chart.
    /// </summary>
    private void PopulateTestData(bool addChart = false)
    {
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? chartObjects = null;
            dynamic? chartObject = null;
            dynamic? chart = null;
            try
            {
                sheet = ctx.Book.Worksheets[1];

                sheet.Range["A1"].Value2 = "Region";
                sheet.Range["B1"].Value2 = "Sales";
                sheet.Range["A2"].Value2 = "North";
                sheet.Range["B2"].Value2 = 45000;
                sheet.Range["A3"].Value2 = "South";
                sheet.Range["B3"].Value2 = 38000;
                sheet.Range["A4"].Value2 = "East";
                sheet.Range["B4"].Value2 = 51000;
                sheet.Range["A5"].Value2 = "West";
                sheet.Range["B5"].Value2 = 42000;

                if (addChart)
                {
                    chartObjects = sheet.ChartObjects();
                    chartObject = chartObjects.Add(150, 100, 400, 250);
                    chart = chartObject.Chart;
                    chart.SetSourceData(sheet.Range["A1:B5"]);
                    chart.ChartType = 51; // xlColumnClustered
                }
            }
            finally
            {
                ComUtilities.Release(ref chart);
                ComUtilities.Release(ref chartObject);
                ComUtilities.Release(ref chartObjects);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private void PopulateColoredBlock(string sheetName, string rangeAddress, int fillColor)
    {
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            try
            {
                sheet = ctx.Book.Worksheets[sheetName];
                range = sheet.Range[rangeAddress];
                range.Value2 = "Screenshot";
                range.Interior.Color = fillColor;
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static void AssertImageContainsNonWhitePixels(string base64)
    {
        using var stream = new MemoryStream(Convert.FromBase64String(base64));
        using var bitmap = new Bitmap(stream);

        for (int y = 0; y < bitmap.Height; y++)
        {
            for (int x = 0; x < bitmap.Width; x++)
            {
                Color pixel = bitmap.GetPixel(x, y);
                if (pixel.R < 245 || pixel.G < 245 || pixel.B < 245)
                {
                    return;
                }
            }
        }

        Assert.Fail("Expected screenshot to contain non-white pixels.");
    }

    [Fact]
    public void CaptureRange_ConsecutiveCalls_AllSucceed()
    {
        // Rapid successive captures must each restore the view cleanly for the next one
        var batch = _fixture.BatchToken;
        PopulateTestData(addChart: true);

        for (int i = 0; i < 3; i++)
        {
            var result = _screenshotCommands.CaptureRange(batch, rangeAddress: "A1:B5");
            Assert.True(result.Success, $"CaptureRange call {i + 1} failed: {result.ErrorMessage}");
            Assert.NotNull(result.ImageBase64);
        }
    }

    [Fact]
    public void CaptureRange_NonActiveSheet_ReturnsNonBlankImage()
    {
        var batch = _fixture.BatchToken;

        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic? newSheet = null;
            dynamic? firstSheet = null;
            try
            {
                newSheet = ctx.Book.Worksheets.Add();
                newSheet.Name = "CaptureTarget";
                firstSheet = ctx.Book.Worksheets[1];
                firstSheet.Activate();
            }
            finally
            {
                ComUtilities.Release(ref firstSheet);
                ComUtilities.Release(ref newSheet);
            }
        });

        PopulateColoredBlock("CaptureTarget", "A1:H12", 255);

        var result = _screenshotCommands.CaptureRange(batch, sheetName: "CaptureTarget", rangeAddress: "A1:H12", quality: ScreenshotQuality.High);

        Assert.True(result.Success, $"CaptureRange failed: {result.ErrorMessage}");
        AssertImageContainsNonWhitePixels(result.ImageBase64);
    }

    [Fact]
    public void CaptureRange_OffscreenRange_ReturnsNonBlankImage()
    {
        var batch = _fixture.BatchToken;

        PopulateColoredBlock("Sheet1", "A200:H220", 65535);

        var result = _screenshotCommands.CaptureRange(batch, rangeAddress: "A200:H220", quality: ScreenshotQuality.High);

        Assert.True(result.Success, $"CaptureRange failed: {result.ErrorMessage}");
        AssertImageContainsNonWhitePixels(result.ImageBase64);
    }
}
