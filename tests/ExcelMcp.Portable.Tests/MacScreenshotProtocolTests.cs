using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacScreenshotProtocolTests
{
    [Fact]
    public void HelperLocator_RequiresExactBundledPath()
    {
        if (!OperatingSystem.IsMacOS())
        {
            return;
        }
        var directory = Directory.CreateTempSubdirectory("excelmcp-capture-locator-");
        try
        {
            Assert.Throws<FileNotFoundException>(() => MacScreenshotHelperLocator.Resolve(directory.FullName));
        }
        finally
        {
            directory.Delete(recursive: true);
        }
    }

    [Fact]
    public void HelperLocator_RejectsNonExecutableFile()
    {
        if (!OperatingSystem.IsMacOS())
        {
            return;
        }
        var directory = CreateHelperDirectory();
        var helper = Path.Combine(directory.FullName, "helpers", "excelmcp-screencapture");
        try
        {
            File.WriteAllBytes(helper, MachOHeaderForCurrentArchitecture());
            File.SetUnixFileMode(helper, UnixFileMode.UserRead | UnixFileMode.UserWrite);

            Assert.Throws<InvalidOperationException>(() => MacScreenshotHelperLocator.Resolve(directory.FullName));
        }
        finally
        {
            directory.Delete(recursive: true);
        }
    }

    [Fact]
    public void HelperLocator_RejectsWrongArchitecture()
    {
        if (!OperatingSystem.IsMacOS())
        {
            return;
        }
        var directory = CreateHelperDirectory();
        var helper = Path.Combine(directory.FullName, "helpers", "excelmcp-screencapture");
        try
        {
            File.WriteAllBytes(helper, MachOHeaderForOtherArchitecture());
            File.SetUnixFileMode(helper, UnixFileMode.UserRead | UnixFileMode.UserExecute);

            Assert.Throws<InvalidOperationException>(() => MacScreenshotHelperLocator.Resolve(directory.FullName));
        }
        finally
        {
            directory.Delete(recursive: true);
        }
    }

    [Theory]
    [InlineData("medium", "jpeg", 0.75)]
    [InlineData("high", "png", 1.0)]
    [InlineData("low", "jpeg", 0.5)]
    public void QualityPolicy_MatchesSharedContract(string quality, string format, double scale)
    {
        var policy = MacScreenshotQualityPolicy.Parse(quality);

        Assert.Equal(format, policy.Format);
        Assert.Equal(scale, policy.Scale);
    }

    [Fact]
    public void CaptureRequest_RejectsCropOutsideCapturedWindow()
    {
        var request = Request(new MacScreenshotPixelRect(900, 700, 200, 100));

        var error = Assert.Throws<ArgumentOutOfRangeException>(() => request.Validate());

        Assert.Equal("CropPixels", error.ParamName);
    }

    [Fact]
    public void CropConverter_SubtractsIndependentlyConvertedWindowEdges()
    {
        var crop = MacScreenshotCropConverter.ConvertCommonTopLeftScreenSpace(
            new MacScreenshotScreenRect(380, 240, 1180, 840),
            new MacScreenshotScreenRect(120, 80, 1720, 980),
            capturedWidth: 1600,
            capturedHeight: 900);

        Assert.Equal(new MacScreenshotPixelRect(260, 160, 800, 600), crop);
    }

    [Fact]
    public void CropConverter_RejectsWindowDimensionDrift()
    {
        var error = Assert.Throws<InvalidOperationException>(() =>
            MacScreenshotCropConverter.ConvertCommonTopLeftScreenSpace(
                new MacScreenshotScreenRect(380, 240, 1180, 840),
                new MacScreenshotScreenRect(120, 80, 1720, 980),
                capturedWidth: 1601,
                capturedHeight: 900));

        Assert.Contains("do not match", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void CropConverter_RejectsRangeOutsideExactWindow()
    {
        var error = Assert.Throws<InvalidOperationException>(() =>
            MacScreenshotCropConverter.ConvertCommonTopLeftScreenSpace(
                new MacScreenshotScreenRect(80, 240, 1180, 840),
                new MacScreenshotScreenRect(120, 80, 1720, 980),
                capturedWidth: 1600,
                capturedHeight: 900));

        Assert.Contains("outside", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PreparedGeometry_ParsesExactOfficeDesktopContract()
    {
        using var document = JsonDocument.Parse(PreparedGeometryJson());

        var result = MacScreenshotPreparedGeometry.Parse(document.RootElement);

        Assert.Equal("capture-token", result.CaptureToken);
        Assert.Equal("file:///exact/owned.xlsx", result.WorkbookUrl);
        Assert.Equal("worksheet-id", result.WorksheetId);
        Assert.Equal("Data", result.WorksheetName);
        Assert.Equal("Data!B2:D8", result.RangeAddress);
        Assert.Equal(42, result.WindowNumber);
        Assert.Equal(new MacScreenshotScreenRect(220, 320, 820, 1120), result.ScreenRect);
        Assert.Equal(
            new MacScreenshotScreenRect(40, 60, 2440, 1860),
            result.ExcelWindowScreenRect);
    }

    [Theory]
    [InlineData("\"left\": 220", "\"left\": 220.5", "integral 'left'")]
    [InlineData("\"width\": 600", "\"width\": 601", "inconsistent edges")]
    [InlineData("\"unit\": \"physicalPixel\"", "\"unit\": \"point\"", "unsupported coordinate")]
    [InlineData(
        "\"origin\": \"topLeftGlobalScreen\"",
        "\"origin\": \"bottomLeftGlobalScreen\"",
        "unsupported coordinate")]
    public void PreparedGeometry_RejectsUnprovenOrInexactCoordinates(
        string original,
        string replacement,
        string expectedMessage)
    {
        var invalid = PreparedGeometryJson().Replace(
            original,
            replacement,
            StringComparison.Ordinal);
        using var document = JsonDocument.Parse(invalid);

        var error = Assert.Throws<InvalidOperationException>(() =>
            MacScreenshotPreparedGeometry.Parse(document.RootElement));

        Assert.Contains(expectedMessage, error.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("\"errorMessage\": null,", "")]
    [InlineData("\"type\": \"workbook\"", "\"type\": \"chart\"")]
    public void PreparedGeometry_RejectsIncompleteOrNonWorkbookEnvelope(
        string original,
        string replacement)
    {
        var invalid = PreparedGeometryJson().Replace(
            original,
            replacement,
            StringComparison.Ordinal);
        using var document = JsonDocument.Parse(invalid);

        Assert.Throws<InvalidOperationException>(() =>
            MacScreenshotPreparedGeometry.Parse(document.RootElement));
    }

    [Fact]
    public void WindowIdentity_CorrelatesExactWorkbookAndOfficeWindow()
    {
        var expectedPath = Path.GetFullPath("owned.xlsx");
        using var document = JsonDocument.Parse(
            $$"""{"success":true,"filePath":"{{JsonEncodedText.Encode(expectedPath)}}","processId":123,"windowId":456,"windowNumber":42}""");

        var identity = MacScreenshotWindowIdentity.Parse(document.RootElement, expectedPath, 42);

        Assert.Equal(123, identity.ProcessId);
        Assert.Equal(456u, identity.WindowId);
        Assert.Equal(42, identity.WindowNumber);
    }

    [Fact]
    public void PreparedGeometry_CreatesOnlyCorrelatedCaptureRequest()
    {
        using var document = JsonDocument.Parse(PreparedGeometryJson());
        var geometry = MacScreenshotPreparedGeometry.Parse(document.RootElement);
        var identity = new MacScreenshotWindowIdentity("/exact/owned.xlsx", 123, 456, 42);

        var request = geometry.CreateCaptureRequest(identity, 2400, 1800, "medium");

        Assert.Equal(new MacScreenshotPixelRect(180, 260, 600, 800), request.CropPixels);
        Assert.Equal(123, request.ProcessId);
        Assert.Equal(456u, request.WindowId);
        Assert.Equal("medium", request.Quality);
        Assert.Throws<InvalidOperationException>(() =>
            geometry.CreateCaptureRequest(identity with { WindowNumber = 43 }, 2400, 1800, "medium"));
    }

    [Fact]
    public async Task Backend_SendsVersionedExactWindowRequestAndReadsImageResult()
    {
        ProcessStartInfo? invoked = null;
        string? input = null;
        var backend = new MacScreenshotBackend(
            "/bundle/ExcelMcp.ScreenCaptureHelper",
            (start, stdin, cancellationToken) =>
            {
                invoked = start;
                input = stdin;
                return Task.FromResult(new MacProcessResult(
                    0,
                    """{"version":1,"success":true,"imageBase64":"aW1hZ2U=","mimeType":"image/png","width":100,"height":50,"error":null}""",
                    ""));
            });

        var result = await backend.CaptureAsync(Request(new MacScreenshotPixelRect(10, 20, 100, 50)),
            TimeSpan.FromSeconds(5));

        Assert.Equal("/bundle/ExcelMcp.ScreenCaptureHelper", invoked?.FileName);
        using var request = JsonDocument.Parse(Assert.IsType<string>(input));
        Assert.Equal(1, request.RootElement.GetProperty("version").GetInt32());
        Assert.Equal(123, request.RootElement.GetProperty("processId").GetInt32());
        Assert.Equal(456u, request.RootElement.GetProperty("windowId").GetUInt32());
        Assert.Equal("high", request.RootElement.GetProperty("quality").GetString());
        Assert.Equal("aW1hZ2U=", result.ImageBase64);
        Assert.Equal("image/png", result.MimeType);
        Assert.Equal(100, result.Width);
        Assert.Equal(50, result.Height);
    }

    [Theory]
    [InlineData("medium", "image/png", 75, 38)]
    [InlineData("medium", "image/jpeg", 100, 50)]
    [InlineData("high", "image/png", 99, 50)]
    [InlineData("low", "image/jpeg", 50, 24)]
    public async Task Backend_RejectsOutputThatDoesNotMatchRequestedQuality(
        string quality,
        string mimeType,
        int width,
        int height)
    {
        var backend = new MacScreenshotBackend(
            "/bundle/ExcelMcp.ScreenCaptureHelper",
            (start, stdin, cancellationToken) => Task.FromResult(new MacProcessResult(
                0,
                $$"""{"version":1,"success":true,"imageBase64":"aW1hZ2U=","mimeType":"{{mimeType}}","width":{{width}},"height":{{height}},"error":null}""",
                "")));
        var request = Request(new MacScreenshotPixelRect(0, 0, 100, 50)) with
        {
            Quality = quality
        };

        await Assert.ThrowsAsync<InvalidOperationException>(() =>
            backend.CaptureAsync(request, TimeSpan.FromSeconds(5)));
    }

    [Fact]
    public async Task Backend_RejectsSuccessfulOutputWithoutExplicitNullError()
    {
        var backend = new MacScreenshotBackend(
            "/bundle/ExcelMcp.ScreenCaptureHelper",
            (start, stdin, cancellationToken) => Task.FromResult(new MacProcessResult(
                0,
                """{"version":1,"success":true,"imageBase64":"aW1hZ2U=","mimeType":"image/png","width":100,"height":50}""",
                "")));

        await Assert.ThrowsAsync<InvalidOperationException>(() =>
            backend.CaptureAsync(
                Request(new MacScreenshotPixelRect(0, 0, 100, 50)),
                TimeSpan.FromSeconds(5)));
    }

    [Fact]
    public async Task Backend_PropagatesPermissionFailureWithoutRequestingConsent()
    {
        var backend = new MacScreenshotBackend(
            "/bundle/ExcelMcp.ScreenCaptureHelper",
            (start, stdin, cancellationToken) => Task.FromResult(new MacProcessResult(
                0,
                """{"version":1,"success":false,"imageBase64":"","mimeType":"","width":0,"height":0,"error":{"code":"screen_recording_required","message":"Screen Recording permission is not ready; no consent was requested."}}""",
                "")));

        var error = await Assert.ThrowsAsync<MacScreenshotException>(() =>
            backend.CaptureAsync(Request(new MacScreenshotPixelRect(0, 0, 100, 50)), TimeSpan.FromSeconds(5)));

        Assert.Equal("screen_recording_required", error.Code);
        Assert.Contains("no consent was requested", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Bridge_ResolvesExactWorkbookWindowIdentityWithoutTitleMatching()
    {
        const string resourceName = "Sbroenne.ExcelMcp.Service.Mac.MacExcelBridge.js";
        using var stream = typeof(MacAutomationHost).Assembly.GetManifestResourceStream(resourceName);
        Assert.NotNull(stream);
        using var reader = new StreamReader(stream);
        var script = reader.ReadToEnd();

        Assert.Contains("command === \"screenshot.window-identity\"", script, StringComparison.Ordinal);
        Assert.Contains("windows[index].windowNumber()", script, StringComparison.Ordinal);
        Assert.Contains("windows[index].id()", script, StringComparison.Ordinal);
        Assert.DoesNotContain(".caption()", script, StringComparison.Ordinal);
        Assert.DoesNotContain(".name()", script[script.IndexOf(
            "function workbookWindowIdentity", StringComparison.Ordinal)..script.IndexOf(
            "function scenarioByName", StringComparison.Ordinal)], StringComparison.Ordinal);
    }

    private static MacScreenshotRequest Request(MacScreenshotPixelRect crop) =>
        new(1, 123, 456, 1000, 800, crop, "high");

    private static DirectoryInfo CreateHelperDirectory()
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-capture-locator-");
        Directory.CreateDirectory(Path.Combine(directory.FullName, "helpers"));
        return directory;
    }

    private static byte[] MachOHeaderForCurrentArchitecture() =>
        MachOHeader(System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture
            == System.Runtime.InteropServices.Architecture.Arm64 ? 0x0100000Cu : 0x01000007u);

    private static byte[] MachOHeaderForOtherArchitecture() =>
        MachOHeader(System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture
            == System.Runtime.InteropServices.Architecture.Arm64 ? 0x01000007u : 0x0100000Cu);

    private static byte[] MachOHeader(uint cpuType)
    {
        var result = new byte[8];
        System.Buffers.Binary.BinaryPrimitives.WriteUInt32LittleEndian(result, 0xFEEDFACF);
        System.Buffers.Binary.BinaryPrimitives.WriteUInt32LittleEndian(result.AsSpan(4), cpuType);
        return result;
    }

    private static string PreparedGeometryJson() =>
        """
        {
          "success": true,
          "errorMessage": null,
          "captureToken": "capture-token",
          "workbookUrl": "file:///exact/owned.xlsx",
          "worksheet": { "id": "worksheet-id", "name": "Data" },
          "range": { "address": "Data!B2:D8" },
          "excelWindow": { "windowNumber": 42, "type": "workbook" },
          "screenRect": {
            "left": 220,
            "top": 320,
            "right": 820,
            "bottom": 1120,
            "width": 600,
            "height": 800,
            "unit": "physicalPixel",
            "origin": "topLeftGlobalScreen"
          },
          "excelWindowScreenRect": {
            "left": 40,
            "top": 60,
            "right": 2440,
            "bottom": 1860,
            "width": 2400,
            "height": 1800,
            "unit": "physicalPixel",
            "origin": "topLeftGlobalScreen"
          }
        }
        """;
}
