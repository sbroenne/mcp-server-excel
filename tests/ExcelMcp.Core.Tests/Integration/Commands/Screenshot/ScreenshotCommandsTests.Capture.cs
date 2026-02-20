// <copyright file="ScreenshotCommandsTests.Capture.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.Screenshot;

/// <summary>
/// Tests for CaptureRange and CaptureSheet operations.
/// These exercise the CopyPicture + ChartObject.Export pipeline including retry logic.
/// </summary>
public partial class ScreenshotCommandsTests
{
    /// <summary>
    /// Helper: populates a test file with sample data and optionally a chart.
    /// </summary>
    private static void PopulateTestData(IExcelBatch batch, bool addChart = false)
    {
        batch.Execute((ctx, ct) =>
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

    [Fact]
    public void CaptureRange_SmallRange_ReturnsValidPng()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(show: true, operationTimeout: null, testFile);
        PopulateTestData(batch);

        // Act — High quality produces PNG
        var result = _commands.CaptureRange(batch, rangeAddress: "A1:B5", quality: ScreenshotQuality.High);

        // Assert
        Assert.True(result.Success, $"CaptureRange failed: {result.ErrorMessage}");
        Assert.NotNull(result.ImageBase64);
        Assert.NotEmpty(result.ImageBase64);
        Assert.Equal("image/png", result.MimeType);
        Assert.True(result.Width > 0, "Width should be positive");
        Assert.True(result.Height > 0, "Height should be positive");

        // Verify it's valid base64 that decodes to a PNG
        byte[] imageBytes = Convert.FromBase64String(result.ImageBase64);
        Assert.True(imageBytes.Length > 100, "Image should be more than 100 bytes");
        // PNG magic bytes: 137 80 78 71
        Assert.Equal(137, imageBytes[0]);
        Assert.Equal(80, imageBytes[1]);
        Assert.Equal(78, imageBytes[2]);
        Assert.Equal(71, imageBytes[3]);
    }

    [Fact]
    public void CaptureRange_MediumQuality_ReturnsJpeg()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(show: true, operationTimeout: null, testFile);
        PopulateTestData(batch);

        // Act — Medium quality (default) produces JPEG
        var result = _commands.CaptureRange(batch, rangeAddress: "A1:B5");

        // Assert
        Assert.True(result.Success, $"CaptureRange failed: {result.ErrorMessage}");
        Assert.NotNull(result.ImageBase64);
        Assert.NotEmpty(result.ImageBase64);
        Assert.Equal("image/jpeg", result.MimeType);
        Assert.True(result.Width > 0, "Width should be positive");
        Assert.True(result.Height > 0, "Height should be positive");

        // Verify it's valid JPEG (SOI marker: FF D8)
        byte[] imageBytes = Convert.FromBase64String(result.ImageBase64);
        Assert.True(imageBytes.Length > 100, "Image should be more than 100 bytes");
        Assert.Equal(0xFF, imageBytes[0]);
        Assert.Equal(0xD8, imageBytes[1]);
    }

    [Fact]
    public void CaptureRange_AreaWithChart_ReturnsLargerImage()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(show: true, operationTimeout: null, testFile);
        PopulateTestData(batch, addChart: true);

        // Act — capture a wider area that includes the chart region (High = PNG for size comparison)
        var result = _commands.CaptureRange(batch, rangeAddress: "A1:M20", quality: ScreenshotQuality.High);

        // Assert
        Assert.True(result.Success, $"CaptureRange failed: {result.ErrorMessage}");
        Assert.NotNull(result.ImageBase64);
        Assert.Equal("image/png", result.MimeType);
        Assert.True(result.Width > 0);
        Assert.True(result.Height > 0);

        byte[] imageBytes = Convert.FromBase64String(result.ImageBase64);
        Assert.True(imageBytes.Length > 500, "Image with chart area should be larger");
    }

    [Fact]
    public void CaptureSheet_NamedSheet_ReturnsValidPng()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(show: true, operationTimeout: null, testFile);
        PopulateTestData(batch);

        // Get the actual sheet name
        string sheetName = batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            return sheet.Name?.ToString() ?? "Sheet1";
        });

        // Act — capture entire used area by sheet name (High = PNG for magic-byte assertion)
        var result = _commands.CaptureSheet(batch, sheetName, quality: ScreenshotQuality.High);

        // Assert
        Assert.True(result.Success, $"CaptureSheet failed: {result.ErrorMessage}");
        Assert.NotNull(result.ImageBase64);
        Assert.NotEmpty(result.ImageBase64);
        Assert.Equal("image/png", result.MimeType);
        Assert.True(result.Width > 0);
        Assert.True(result.Height > 0);
        Assert.Equal(sheetName, result.SheetName);
    }

    [Fact]
    public void CaptureSheet_ActiveSheet_ReturnsValidPng()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(show: true, operationTimeout: null, testFile);
        PopulateTestData(batch);

        // Act — capture active sheet (no sheetName specified), High quality for PNG assertion
        var result = _commands.CaptureSheet(batch, quality: ScreenshotQuality.High);

        // Assert
        Assert.True(result.Success, $"CaptureSheet failed: {result.ErrorMessage}");
        Assert.NotNull(result.ImageBase64);
        Assert.Equal("image/png", result.MimeType);
    }

    [Fact]
    public void CaptureRange_DefaultRange_ReturnsValidPng()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(show: true, operationTimeout: null, testFile);
        PopulateTestData(batch);

        // Act — use default range (A1:Z30) with no sheetName, High quality for PNG assertion
        var result = _commands.CaptureRange(batch, quality: ScreenshotQuality.High);

        // Assert
        Assert.True(result.Success, $"CaptureRange failed: {result.ErrorMessage}");
        Assert.NotNull(result.ImageBase64);
        Assert.Equal("image/png", result.MimeType);
        Assert.True(result.Width > 0);
        Assert.True(result.Height > 0);
    }

    [Fact]
    public void CaptureRange_MessageIncludesDimensions()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(show: true, operationTimeout: null, testFile);
        PopulateTestData(batch);

        // Act
        var result = _commands.CaptureRange(batch, rangeAddress: "A1:B5");

        // Assert — message should contain pixel dimensions
        Assert.True(result.Success);
        Assert.Contains("px", result.Message);
    }

    [Fact]
    public void CaptureRange_ConsecutiveCalls_AllSucceed()
    {
        // This test validates the retry logic handles rapid successive CopyPicture calls
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(show: true, operationTimeout: null, testFile);
        PopulateTestData(batch, addChart: true);

        for (int i = 0; i < 3; i++)
        {
            var result = _commands.CaptureRange(batch, rangeAddress: "A1:B5");
            Assert.True(result.Success, $"CaptureRange call {i + 1} failed: {result.ErrorMessage}");
            Assert.NotNull(result.ImageBase64);
        }
    }
}
