// <copyright file="ScreenshotCommandsTests.Capture.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

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
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? source = null;
            Excel.ChartObjects? chartObjects = null;
            Excel.ChartObject? chartObject = null;
            Excel.Chart? chart = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                source = sheet.Range["A1:B5"];
                source.Value2 = new object[,]
                {
                    { "Region", "Sales" }, { "North", 45000 }, { "South", 38000 },
                    { "East", 51000 }, { "West", 42000 }
                };

                if (addChart)
                {
                    chartObjects = (Excel.ChartObjects)sheet.ChartObjects();
                    chartObject = chartObjects.Add(150, 100, 400, 250);
                    chart = chartObject.Chart;
                    chart.SetSourceData(source);
                    chart.ChartType = Excel.XlChartType.xlColumnClustered;
                }
            }
            finally
            {
                ComUtilities.Release(ref chart);
                ComUtilities.Release(ref chartObject);
                ComUtilities.Release(ref chartObjects);
                ComUtilities.Release(ref source);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    private void PopulateColoredBlock(string sheetName, string rangeAddress, int fillColor)
    {
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.Interior? interior = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                range = sheet.Range[rangeAddress];
                range.Value2 = "Screenshot";
                interior = range.Interior;
                interior.Color = fillColor;
                Assert.Equal(fillColor, Convert.ToInt32(interior.Color, System.Globalization.CultureInfo.InvariantCulture));
                Assert.All(Assert.IsType<object[,]>((object?)range.Value2).Cast<object>(),
                    value => Assert.Equal("Screenshot", value));
            }
            finally
            {
                ComUtilities.Release(ref interior);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    private static void AssertImageContainsMarker(ScreenshotResult result, Color marker)
    {
        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(result.ErrorMessage));
        Assert.NotEmpty(result.ImageBase64);
        using var stream = new MemoryStream(Convert.FromBase64String(result.ImageBase64));
        using var bitmap = new Bitmap(stream);
        Assert.Equal(result.Width, bitmap.Width);
        Assert.Equal(result.Height, bitmap.Height);
        var matching = 0;
        for (int y = 0; y < bitmap.Height; y++)
        {
            for (int x = 0; x < bitmap.Width; x++)
            {
                Color pixel = bitmap.GetPixel(x, y);
                if (Math.Abs(pixel.R - marker.R) < 25 &&
                    Math.Abs(pixel.G - marker.G) < 25 &&
                    Math.Abs(pixel.B - marker.B) < 25)
                {
                    matching++;
                }
            }
        }

        Assert.True(matching >= 10, $"Expected screenshot to contain its {marker} marker.");
    }

    private void AssertColoredBlockPreserved(string sheetName, string rangeAddress, int fillColor, bool protectedSheet = false)
    {
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.Interior? interior = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                range = sheet.Range[rangeAddress];
                interior = range.Interior;
                Assert.Equal(protectedSheet, sheet.ProtectContents);
                Assert.Equal(fillColor, Convert.ToInt32(interior.Color, System.Globalization.CultureInfo.InvariantCulture));
                Assert.All(Assert.IsType<object[,]>((object?)range.Value2).Cast<object>(),
                    value => Assert.Equal("Screenshot", value));
            }
            finally
            {
                ComUtilities.Release(ref interior);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    private string ReadActiveSheetName() =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            try
            {
                sheet = (Excel.Worksheet)context.Book.ActiveSheet;
                return sheet.Name;
            }
            finally { ComUtilities.Release(ref sheet); }
        });

    [Fact]
    public void CaptureRange_ConsecutiveCalls_AllSucceed()
    {
        // Rapid successive captures must each restore the view cleanly for the next one
        var batch = _fixture.BatchToken;
        PopulateTestData(addChart: true);
        PopulateColoredBlock("Sheet1", "C1:D5", 255);

        for (int i = 0; i < 3; i++)
        {
            var result = _screenshotCommands.CaptureRange(batch, rangeAddress: "A1:D5");
            AssertImageContainsMarker(result, Color.Red);
            AssertColoredBlockPreserved("Sheet1", "C1:D5", 255);
        }
    }

    [Fact]
    public void CaptureRange_NonActiveSheet_ReturnsNonBlankImage()
    {
        var batch = _fixture.BatchToken;

        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? newSheet = null;
            Excel.Worksheet? firstSheet = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                firstSheet = (Excel.Worksheet)sheets[1];
                newSheet = (Excel.Worksheet)sheets.Add();
                newSheet.Name = "CaptureTarget";
                firstSheet.Activate();
            }
            finally
            {
                ComUtilities.Release(ref firstSheet);
                ComUtilities.Release(ref newSheet);
                ComUtilities.Release(ref sheets);
            }
        });

        PopulateColoredBlock("CaptureTarget", "A1:H12", 255);
        Assert.Equal("Sheet1", ReadActiveSheetName());

        var result = _screenshotCommands.CaptureRange(batch, sheetName: "CaptureTarget", rangeAddress: "A1:H12", quality: ScreenshotQuality.High);

        Assert.True(result.Success, $"CaptureRange failed: {result.ErrorMessage}");
        AssertImageContainsMarker(result, Color.Red);
        Assert.Equal("CaptureTarget", result.SheetName);
        Assert.Equal("$A$1:$H$12", result.RangeAddress);
        AssertColoredBlockPreserved("CaptureTarget", "A1:H12", 255);
        Assert.Equal("Sheet1", ReadActiveSheetName());
    }

    [Fact]
    public void CaptureRange_OffscreenRange_ReturnsNonBlankImage()
    {
        var batch = _fixture.BatchToken;

        PopulateColoredBlock("Sheet1", "A200:H220", 65535);

        var result = _screenshotCommands.CaptureRange(batch, rangeAddress: "A200:H220", quality: ScreenshotQuality.High);

        Assert.True(result.Success, $"CaptureRange failed: {result.ErrorMessage}");
        AssertImageContainsMarker(result, Color.Yellow);
        Assert.Equal("Sheet1", result.SheetName);
        Assert.Equal("$A$200:$H$220", result.RangeAddress);
        AssertColoredBlockPreserved("Sheet1", "A200:H220", 65535);
    }
}
