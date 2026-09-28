// <copyright file="ScreenshotCommandsTests.ProtectedSheet.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Regression coverage for issue #777: capture failed with COMException 0x800A03EC on a
/// protected worksheet because the old pipeline inserted a temporary ChartObject into the
/// target sheet, which Excel refuses while the sheet is protected.
/// </summary>
public sealed partial class IsolatedServiceScreenshotTests
{
    [Fact]
    public void CaptureRange_ProtectedSheet_ReturnsImage()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        PopulateColoredBlock(sheetName, "A1:D8", 255);
        ProtectSheet(batch, sheetName);

        var result = _screenshotCommands.CaptureRange(batch, sheetName, "A1:D8", ScreenshotQuality.High);

        Assert.True(result.Success, $"CaptureRange failed on a protected sheet: {result.ErrorMessage}");
        AssertImageContainsNonWhitePixels(result.ImageBase64);
    }

    [Fact]
    public void CaptureSheet_ProtectedSheet_ReturnsImage()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        PopulateColoredBlock(sheetName, "A1:D8", 65535);
        ProtectSheet(batch, sheetName);

        var result = _screenshotCommands.CaptureSheet(batch, sheetName, ScreenshotQuality.High);

        Assert.True(result.Success, $"CaptureSheet failed on a protected sheet: {result.ErrorMessage}");
        AssertImageContainsNonWhitePixels(result.ImageBase64);
    }

    [Fact]
    public void CaptureRange_ProtectedSheet_LeavesProtectionIntact()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        PopulateColoredBlock(sheetName, "A1:D8", 255);
        ProtectSheet(batch, sheetName);

        var result = _screenshotCommands.CaptureRange(batch, sheetName, "A1:D8");

        Assert.True(result.Success, $"CaptureRange failed on a protected sheet: {result.ErrorMessage}");

        var protection = _sheetCommands.GetProtection(batch, sheetName);

        Assert.True(protection.Success);
        Assert.True(protection.IsProtected, "Capture must not unprotect the worksheet.");
    }

    private void ProtectSheet(IExcelBatch batch, string sheetName)
    {
        var result = _sheetCommands.SetProtection(batch, sheetName, isProtected: true);
        Assert.True(result.Success, $"Failed to protect '{sheetName}': {result.ErrorMessage}");
    }
}
