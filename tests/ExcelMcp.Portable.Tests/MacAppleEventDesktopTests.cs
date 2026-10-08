using System.Diagnostics;
using System.Text.Json;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

[Trait("RequiresExcel", "true")]
[Collection("Mac Excel E2E")]
public sealed class MacAppleEventDesktopTests
{
    [Fact]
    public async Task UnsupportedNativeVariantsPreservePlatformErrorCategory()
    {
        Assert.True(OperatingSystem.IsMacOS(), "This acceptance test requires macOS desktop Excel.");
        Assert.Equal(0, MacAutomationAccess.Check());
        var start = new ProcessStartInfo(Environment.ProcessPath!)
        {
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };
        start.ArgumentList.Add(Path.Combine(MacExcelE2ETests.FindRepository(),
            "src/ExcelMcp.CLI/bin/Release/net10.0/excelcli.dll"));
        start.ArgumentList.Add(MacAutomationHost.Marker);
        start.ArgumentList.Add("calculation.calculate");
        start.ArgumentList.Add(Environment.ProcessId.ToString(System.Globalization.CultureInfo.InvariantCulture));
        start.ArgumentList.Add(TimeSpan.FromSeconds(10).Ticks.ToString(System.Globalization.CultureInfo.InvariantCulture));
        using var process = Process.Start(start) ?? throw new InvalidOperationException("Could not start the native calculation child.");
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(15));
        var stdout = process.StandardOutput.ReadToEndAsync(deadline.Token);
        var stderr = process.StandardError.ReadToEndAsync(deadline.Token);
        try
        {
            await process.StandardInput.WriteAsync(
                """{"filePath":"opaque.xlsx","scope":"application","sheetName":""}""".AsMemory(), deadline.Token);
            process.StandardInput.Close();
            await process.WaitForExitAsync(deadline.Token);
            Assert.Equal(0, process.ExitCode);
            using var response = JsonDocument.Parse(await stdout);
            Assert.False(response.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal("PlatformNotSupported", response.RootElement.GetProperty("errorCategory").GetString());
            Assert.Contains("no calculation was attempted", response.RootElement.GetProperty("errorMessage").GetString(),
                StringComparison.Ordinal);
            Assert.Equal("", await stderr);
        }
        finally
        {
            if (!process.HasExited)
            {
                process.Kill(entireProcessTree: true);
                await process.WaitForExitAsync();
            }
        }
    }

    [Fact]
    public void MissingMacroCannotBeReportedAsAnInstalledCompatibleHelper()
    {
        Assert.True(OperatingSystem.IsMacOS(), "This acceptance test requires macOS desktop Excel.");
        Assert.Equal(0, MacAutomationAccess.Check());
        var error = Assert.ThrowsAny<InvalidOperationException>(() =>
            MacHelperProtocol.ValidateResponse(
                MacNativeHelper.Call($"'ExcelMcpMissing-{Guid.NewGuid():N}.xlam'!ExcelMcpHelper_Info", TimeSpan.FromSeconds(10)),
                MacHelperProtocol.InfoPrimitives));
        if (error is MacExcelOperationException helperError)
        {
            Assert.Equal("HelperMissing", helperError.ErrorCategory);
            Assert.Contains("Install and enable", error.Message, StringComparison.Ordinal);
        }
        else
        {
            Assert.Contains("Excel reply", error.Message, StringComparison.Ordinal);
            Assert.Contains("OSStatus", error.Message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void NativeEventsReadWorkbookNamesWithoutScriptsOrMutation()
    {
        Assert.True(OperatingSystem.IsMacOS(), "This acceptance test requires macOS desktop Excel.");
        Assert.Equal(0, MacAutomationAccess.Check());

        using var application = MacAppleEvents.Create(MacAppleEvents.Code("null"), []);
        var count = MacNativeRange.Count(application, MacExcelDictionary.WorkbookClass, TimeSpan.FromSeconds(10));
        var result = MacNativeWorkbook.ReadNames(false, TimeSpan.FromSeconds(10));
        Assert.Equal(count, result.Length);
        Assert.All(result, item => Assert.False(string.IsNullOrEmpty(item)));
    }

    [Fact]
    public void NativeErrorsPreserveTheirStatusInsteadOfReturningSuccess()
    {
        Assert.True(OperatingSystem.IsMacOS(), "This acceptance test requires macOS desktop Excel.");
        Assert.Equal(0, MacAutomationAccess.Check());

        using var application = MacAppleEvents.Create(MacAppleEvents.Code("null"), []);
        using var workbookName = MacAppleEvents.Text($"ExcelMcpMissing-{Guid.NewGuid():N}.xlsx");
        using var workbook = MacAppleEvents.Object(
            MacAppleEvents.Code("X141"), application, MacAppleEvents.Code("name"), workbookName);
        using var unknown = MacAppleEvents.Property(workbook, MacAppleEvents.Code("pnam"));
        using var appleEvent = MacAppleEvents.Event(MacAppleEvents.Code("core"), MacAppleEvents.Code("getd"));
        MacAppleEvents.Put(appleEvent, MacAppleEvents.Code("----"), unknown);

        var error = Assert.Throws<MacAppleEventException>(() =>
            MacAppleEvents.Send(appleEvent, TimeSpan.FromSeconds(10)));
        Assert.NotEqual(0, error.Status);
        Assert.True(error.IsExcelReply);
        Assert.True(error.Message.Contains("OSStatus", StringComparison.Ordinal), error.ToString());
    }
}
