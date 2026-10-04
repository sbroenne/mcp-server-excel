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
        AssertImageContainsMarker(result, System.Drawing.Color.Red);
        Assert.Equal(sheetName, result.SheetName);
        Assert.Equal("$A$1:$D$8", result.RangeAddress);
        AssertColoredBlockPreserved(sheetName, "A1:D8", 255, protectedSheet: true);
        var protection = RequireSuccess(_sheetCommands.GetProtection(batch, sheetName));
        Assert.True(protection.IsProtected, "Capture must not unprotect the worksheet.");
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
        AssertImageContainsMarker(result, System.Drawing.Color.Yellow);
        Assert.Equal(sheetName, result.SheetName);
        AssertColoredBlockPreserved(sheetName, "A1:D8", 65535, protectedSheet: true);
    }

    private void ProtectSheet(IExcelBatch batch, string sheetName)
    {
        var result = _sheetCommands.SetProtection(batch, sheetName, isProtected: true);
        Assert.True(result.Success, $"Failed to protect '{sheetName}': {result.ErrorMessage}");
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void CaptureRange_RejectedTarget_PreservesProtectedSheetAndAllowsRecovery(bool missingSheet)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        PopulateColoredBlock(sheetName, "A1:D8", 255);
        ProtectSheet(batch, sheetName);
        AssertColoredBlockPreserved(sheetName, "A1:D8", 255, protectedSheet: true);
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _screenshotCommands.CaptureRange(batch, missingSheet ? "MissingCaptureSheet" : sheetName,
                missingSheet ? "A1:D8" : "NotAnExistingRange", ScreenshotQuality.High));
        Assert.Contains("screenshot.capture", exception.Message, StringComparison.Ordinal);
        AssertColoredBlockPreserved(sheetName, "A1:D8", 255, protectedSheet: true);

        var recovered = _screenshotCommands.CaptureRange(batch, sheetName, "A1:D8", ScreenshotQuality.High);
        AssertImageContainsMarker(recovered, System.Drawing.Color.Red);
        AssertColoredBlockPreserved(sheetName, "A1:D8", 255, protectedSheet: true);
    }
}
