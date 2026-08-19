// <copyright file="ScreenshotCommandsTests.ProtectedSheet.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.Screenshot;

/// <summary>
/// Regression coverage for issue #777: capture failed with COMException 0x800A03EC on a
/// protected worksheet because the old pipeline inserted a temporary ChartObject into the
/// target sheet, which Excel refuses while the sheet is protected.
/// </summary>
public partial class ScreenshotCommandsTests
{
    [Fact]
    public void CaptureRange_ProtectedSheet_ReturnsImage()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(show: true, operationTimeout: null, testFile);

        PopulateColoredBlock(batch, "Sheet1", "A1:D8", 255);
        ProtectSheet(batch, "Sheet1");

        var result = _commands.CaptureRange(batch, sheetName: "Sheet1", rangeAddress: "A1:D8", quality: ScreenshotQuality.High);

        Assert.True(result.Success, $"CaptureRange failed on a protected sheet: {result.ErrorMessage}");
        AssertImageContainsNonWhitePixels(result.ImageBase64);
    }

    [Fact]
    public void CaptureSheet_ProtectedSheet_ReturnsImage()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(show: true, operationTimeout: null, testFile);

        PopulateColoredBlock(batch, "Sheet1", "A1:D8", 65535);
        ProtectSheet(batch, "Sheet1");

        var result = _commands.CaptureSheet(batch, sheetName: "Sheet1", quality: ScreenshotQuality.High);

        Assert.True(result.Success, $"CaptureSheet failed on a protected sheet: {result.ErrorMessage}");
        AssertImageContainsNonWhitePixels(result.ImageBase64);
    }

    [Fact]
    public void CaptureRange_ProtectedSheet_LeavesProtectionIntact()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(show: true, operationTimeout: null, testFile);

        PopulateColoredBlock(batch, "Sheet1", "A1:D8", 255);
        ProtectSheet(batch, "Sheet1");

        var result = _commands.CaptureRange(batch, sheetName: "Sheet1", rangeAddress: "A1:D8");

        Assert.True(result.Success, $"CaptureRange failed on a protected sheet: {result.ErrorMessage}");

        var sheetCommands = new SheetCommands();
        var protection = sheetCommands.GetProtection(batch, "Sheet1");

        Assert.True(protection.Success);
        Assert.True(protection.IsProtected, "Capture must not unprotect the worksheet.");
    }

    private static void ProtectSheet(IExcelBatch batch, string sheetName)
    {
        var sheetCommands = new SheetCommands();
        var result = sheetCommands.SetProtection(batch, sheetName, isProtected: true);
        Assert.True(result.Success, $"Failed to protect '{sheetName}': {result.ErrorMessage}");
    }
}
