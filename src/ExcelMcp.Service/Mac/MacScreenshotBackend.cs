using System.Diagnostics;
using System.Buffers.Binary;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal readonly record struct MacScreenshotPixelRect(int X, int Y, int Width, int Height);

internal readonly record struct MacScreenshotScreenRect(int Left, int Top, int Right, int Bottom)
{
    public int Width => checked(Right - Left);
    public int Height => checked(Bottom - Top);

    public void Validate(string parameterName)
    {
        if (Right <= Left || Bottom <= Top)
        {
            throw new ArgumentOutOfRangeException(
                parameterName, "Screenshot screen rectangle must have positive width and height.");
        }
    }
}

internal sealed record MacScreenshotPreparedGeometry(
    string CaptureToken,
    string WorkbookUrl,
    string WorksheetId,
    string WorksheetName,
    string RangeAddress,
    int WindowNumber,
    MacScreenshotScreenRect ScreenRect,
    MacScreenshotScreenRect ExcelWindowScreenRect)
{
    public static MacScreenshotPreparedGeometry Parse(JsonElement value)
    {
        if (value.ValueKind != JsonValueKind.Object
            || !value.TryGetProperty("success", out var success)
            || success.ValueKind is not JsonValueKind.True
            || !value.TryGetProperty("errorMessage", out var errorMessage)
            || errorMessage.ValueKind is not JsonValueKind.Null)
        {
            throw new InvalidOperationException(
                "Office.js screenshot preparation did not return a successful result.");
        }

        var worksheet = RequireObject(value, "worksheet");
        var range = RequireObject(value, "range");
        var excelWindow = RequireObject(value, "excelWindow");
        if (!string.Equals(
                RequireString(excelWindow, "type"),
                "workbook",
                StringComparison.Ordinal))
        {
            throw new InvalidOperationException(
                "Office.js screenshot preparation did not return a workbook window.");
        }
        var screenRect = ParseScreenRect(value, "screenRect");
        var excelWindowScreenRect = ParseScreenRect(value, "excelWindowScreenRect");
        return new MacScreenshotPreparedGeometry(
            RequireString(value, "captureToken"),
            RequireString(value, "workbookUrl"),
            RequireString(worksheet, "id"),
            RequireString(worksheet, "name"),
            RequireString(range, "address"),
            RequirePositiveInt32(excelWindow, "windowNumber"),
            screenRect,
            excelWindowScreenRect);
    }

    public MacScreenshotRequest CreateCaptureRequest(
        MacScreenshotWindowIdentity identity,
        int capturedWidth,
        int capturedHeight,
        string quality)
    {
        ArgumentNullException.ThrowIfNull(identity);
        if (identity.WindowNumber != WindowNumber)
        {
            throw new InvalidOperationException(
                "The native Excel window identity does not match the Office.js prepared window.");
        }
        var crop = MacScreenshotCropConverter.ConvertCommonTopLeftScreenSpace(
            ScreenRect,
            ExcelWindowScreenRect,
            capturedWidth,
            capturedHeight);
        var request = new MacScreenshotRequest(
            1,
            identity.ProcessId,
            identity.WindowId,
            capturedWidth,
            capturedHeight,
            crop,
            quality);
        request.Validate();
        return request;
    }

    private static MacScreenshotScreenRect ParseScreenRect(JsonElement parent, string propertyName)
    {
        var element = RequireObject(parent, propertyName);
        if (!string.Equals(
                RequireString(element, "unit"),
                "physicalPixel",
                StringComparison.Ordinal)
            || !string.Equals(
                RequireString(element, "origin"),
                "topLeftGlobalScreen",
                StringComparison.Ordinal))
        {
            throw new InvalidOperationException(
                $"Office.js screenshot geometry '{propertyName}' uses an unsupported coordinate system.");
        }
        var result = new MacScreenshotScreenRect(
            RequireInt32(element, "left"),
            RequireInt32(element, "top"),
            RequireInt32(element, "right"),
            RequireInt32(element, "bottom"));
        result.Validate(propertyName);
        if (RequirePositiveInt32(element, "width") != result.Width
            || RequirePositiveInt32(element, "height") != result.Height)
        {
            throw new InvalidOperationException(
                $"Office.js screenshot geometry '{propertyName}' contains inconsistent edges and dimensions.");
        }
        return result;
    }

    private static JsonElement RequireObject(JsonElement parent, string propertyName)
    {
        if (!parent.TryGetProperty(propertyName, out var value)
            || value.ValueKind != JsonValueKind.Object)
        {
            throw new InvalidOperationException(
                $"Office.js screenshot preparation is missing object '{propertyName}'.");
        }
        return value;
    }

    private static string RequireString(JsonElement parent, string propertyName)
    {
        if (!parent.TryGetProperty(propertyName, out var value)
            || value.ValueKind != JsonValueKind.String
            || string.IsNullOrWhiteSpace(value.GetString()))
        {
            throw new InvalidOperationException(
                $"Office.js screenshot preparation is missing string '{propertyName}'.");
        }
        return value.GetString()!;
    }

    private static int RequireInt32(JsonElement parent, string propertyName)
    {
        if (!parent.TryGetProperty(propertyName, out var value)
            || value.ValueKind != JsonValueKind.Number
            || !value.TryGetInt32(out var result))
        {
            throw new InvalidOperationException(
                $"Office.js screenshot preparation requires integral '{propertyName}'.");
        }
        return result;
    }

    private static int RequirePositiveInt32(JsonElement parent, string propertyName)
    {
        var result = RequireInt32(parent, propertyName);
        if (result <= 0)
        {
            throw new InvalidOperationException(
                $"Office.js screenshot preparation requires positive '{propertyName}'.");
        }
        return result;
    }
}

internal sealed record MacScreenshotWindowIdentity(
    string FilePath,
    int ProcessId,
    uint WindowId,
    int WindowNumber)
{
    public static MacScreenshotWindowIdentity Parse(
        JsonElement value,
        string expectedFilePath,
        int expectedWindowNumber)
    {
        if (value.ValueKind != JsonValueKind.Object
            || !value.TryGetProperty("success", out var success)
            || success.ValueKind is not JsonValueKind.True)
        {
            throw new InvalidOperationException(
                "Excel did not return a successful native window identity.");
        }
        var filePath = RequireString(value, "filePath");
        var processId = RequirePositiveInt32(value, "processId");
        var windowNumber = RequirePositiveInt32(value, "windowNumber");
        if (!value.TryGetProperty("windowId", out var windowIdElement)
            || windowIdElement.ValueKind != JsonValueKind.Number
            || !windowIdElement.TryGetUInt32(out var windowId)
            || windowId == 0)
        {
            throw new InvalidOperationException(
                "Excel returned an invalid native window identifier.");
        }
        if (!string.Equals(
                MacPathCanonicalizer.Normalize(filePath),
                MacPathCanonicalizer.Normalize(expectedFilePath),
                StringComparison.Ordinal))
        {
            throw new InvalidOperationException(
                "The native Excel window identity belongs to a different workbook.");
        }
        if (windowNumber != expectedWindowNumber)
        {
            throw new InvalidOperationException(
                "The native Excel window identity does not match the Office.js prepared window.");
        }
        return new MacScreenshotWindowIdentity(filePath, processId, windowId, windowNumber);
    }

    private static string RequireString(JsonElement parent, string propertyName)
    {
        if (!parent.TryGetProperty(propertyName, out var value)
            || value.ValueKind != JsonValueKind.String
            || string.IsNullOrWhiteSpace(value.GetString()))
        {
            throw new InvalidOperationException(
                $"Native Excel window identity is missing string '{propertyName}'.");
        }
        return value.GetString()!;
    }

    private static int RequirePositiveInt32(JsonElement parent, string propertyName)
    {
        if (!parent.TryGetProperty(propertyName, out var value)
            || value.ValueKind != JsonValueKind.Number
            || !value.TryGetInt32(out var result)
            || result <= 0)
        {
            throw new InvalidOperationException(
                $"Native Excel window identity requires positive '{propertyName}'.");
        }
        return result;
    }
}

internal static class MacScreenshotCropConverter
{
    public static MacScreenshotPixelRect ConvertCommonTopLeftScreenSpace(
        MacScreenshotScreenRect rangeScreenRect,
        MacScreenshotScreenRect excelWindowScreenRect,
        int capturedWidth,
        int capturedHeight)
    {
        rangeScreenRect.Validate(nameof(rangeScreenRect));
        excelWindowScreenRect.Validate(nameof(excelWindowScreenRect));
        if (capturedWidth != excelWindowScreenRect.Width
            || capturedHeight != excelWindowScreenRect.Height)
        {
            throw new InvalidOperationException(
                "ScreenCaptureKit window dimensions do not match the Excel-provided physical-pixel window rectangle.");
        }

        var crop = new MacScreenshotPixelRect(
            checked(rangeScreenRect.Left - excelWindowScreenRect.Left),
            checked(rangeScreenRect.Top - excelWindowScreenRect.Top),
            rangeScreenRect.Width,
            rangeScreenRect.Height);
        if (crop.X < 0
            || crop.Y < 0
            || (long)crop.X + crop.Width > capturedWidth
            || (long)crop.Y + crop.Height > capturedHeight)
        {
            throw new InvalidOperationException(
                "Excel-provided range geometry falls outside the exact captured workbook window.");
        }
        return crop;
    }
}

internal sealed record MacScreenshotRequest(
    int Version,
    int ProcessId,
    uint WindowId,
    int CapturedWidth,
    int CapturedHeight,
    MacScreenshotPixelRect CropPixels,
    string Quality)
{
    public void Validate()
    {
        if (Version != 1)
        {
            throw new ArgumentOutOfRangeException(nameof(Version), "Screenshot protocol version must be 1.");
        }
        if (ProcessId <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(ProcessId), "Excel process ID must be positive.");
        }
        if (WindowId == 0)
        {
            throw new ArgumentOutOfRangeException(nameof(WindowId), "Excel window ID must be positive.");
        }
        if (CapturedWidth <= 0 || CapturedHeight <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(CapturedWidth), "Captured window dimensions must be positive.");
        }
        _ = MacScreenshotQualityPolicy.Parse(Quality);
        if (CropPixels.X < 0
            || CropPixels.Y < 0
            || CropPixels.Width <= 0
            || CropPixels.Height <= 0
            || (long)CropPixels.X + CropPixels.Width > CapturedWidth
            || (long)CropPixels.Y + CropPixels.Height > CapturedHeight)
        {
            throw new ArgumentOutOfRangeException(
                nameof(CropPixels), "Screenshot crop must be a positive rectangle inside the captured Excel window.");
        }
    }
}

internal readonly record struct MacScreenshotQualityPolicy(string Format, double Scale)
{
    public static MacScreenshotQualityPolicy Parse(string quality) =>
        quality?.ToLowerInvariant() switch
        {
            "medium" => new("jpeg", 0.75),
            "high" => new("png", 1.0),
            "low" => new("jpeg", 0.5),
            _ => throw new ArgumentOutOfRangeException(
                nameof(quality), quality, "Screenshot quality must be medium, high, or low.")
        };
}

internal sealed record MacScreenshotCaptureResult(
    string ImageBase64,
    string MimeType,
    int Width,
    int Height);

internal sealed class MacScreenshotBackend
{
    private readonly string _helperPath;
    private readonly Func<ProcessStartInfo, string?, CancellationToken, Task<MacProcessResult>> _runProcess;

    public MacScreenshotBackend(
        string helperPath,
        Func<ProcessStartInfo, string?, CancellationToken, Task<MacProcessResult>>? runProcess = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(helperPath);
        _helperPath = helperPath;
        _runProcess = runProcess ?? RunProcessAsync;
    }

    public async Task<MacScreenshotCaptureResult> CaptureAsync(
        MacScreenshotRequest request,
        TimeSpan timeout)
    {
        ArgumentNullException.ThrowIfNull(request);
        request.Validate();
        using var timeoutCts = new CancellationTokenSource(timeout);
        var start = new ProcessStartInfo(_helperPath)
        {
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        var input = JsonSerializer.Serialize(request, ServiceProtocol.JsonOptions);

        MacProcessResult processResult;
        try
        {
            processResult = await _runProcess(start, input, timeoutCts.Token);
        }
        catch (OperationCanceledException) when (timeoutCts.IsCancellationRequested)
        {
            throw new TimeoutException(
                $"Mac screenshot helper exceeded {timeout.TotalSeconds:0.###} seconds.");
        }

        if (processResult.ExitCode != 0)
        {
            throw new InvalidOperationException(
                $"Mac screenshot helper failed with exit code {processResult.ExitCode}: " +
                SanitizeError(processResult.StandardError));
        }

        using var document = JsonDocument.Parse(processResult.StandardOutput);
        var root = document.RootElement;
        if (root.GetProperty("version").GetInt32() != 1)
        {
            throw new InvalidOperationException("Mac screenshot helper returned an unsupported protocol version.");
        }
        if (!root.GetProperty("success").GetBoolean())
        {
            var error = root.GetProperty("error");
            throw new MacScreenshotException(
                error.GetProperty("code").GetString() ?? "capture_failed",
                error.GetProperty("message").GetString() ?? "Mac screenshot capture failed.");
        }
        if (root.ValueKind != JsonValueKind.Object
            || !root.TryGetProperty("error", out var successError)
            || successError.ValueKind is not JsonValueKind.Null)
        {
            throw new InvalidOperationException("Successful Mac screenshot response contained an error.");
        }

        var imageBase64 = root.GetProperty("imageBase64").GetString() ?? string.Empty;
        if (imageBase64.Length == 0 || Convert.FromBase64String(imageBase64).Length == 0)
        {
            throw new InvalidOperationException("Mac screenshot helper returned empty image data.");
        }
        var mimeType = root.GetProperty("mimeType").GetString() ?? string.Empty;
        var width = root.GetProperty("width").GetInt32();
        var height = root.GetProperty("height").GetInt32();
        var quality = MacScreenshotQualityPolicy.Parse(request.Quality);
        var expectedMimeType = quality.Format == "png" ? "image/png" : "image/jpeg";
        var expectedWidth = Math.Max(1, (int)Math.Round(
            request.CropPixels.Width * quality.Scale,
            MidpointRounding.AwayFromZero));
        var expectedHeight = Math.Max(1, (int)Math.Round(
            request.CropPixels.Height * quality.Scale,
            MidpointRounding.AwayFromZero));
        if (mimeType != expectedMimeType || width != expectedWidth || height != expectedHeight)
        {
            throw new InvalidOperationException(
                "Mac screenshot helper returned image metadata that does not match the requested crop and quality.");
        }
        return new MacScreenshotCaptureResult(imageBase64, mimeType, width, height);
    }

    private static async Task<MacProcessResult> RunProcessAsync(
        ProcessStartInfo start,
        string? input,
        CancellationToken cancellationToken)
    {
        using var process = new Process { StartInfo = start };
        cancellationToken.ThrowIfCancellationRequested();
        if (!process.Start())
        {
            throw new InvalidOperationException($"Could not start '{start.FileName}'.");
        }
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        try
        {
            if (input is not null)
            {
                await process.StandardInput.WriteAsync(input.AsMemory(), cancellationToken);
                process.StandardInput.Close();
            }
            await process.WaitForExitAsync(cancellationToken);
        }
        catch (OperationCanceledException)
        {
            if (!process.HasExited)
            {
                process.Kill();
            }
            await process.WaitForExitAsync();
            await Task.WhenAll(stdout, stderr);
            throw;
        }
        return new MacProcessResult(process.ExitCode, await stdout, await stderr);
    }

    private static string SanitizeError(string error)
    {
        var text = error.Trim();
        return text.Length > 500 ? text[..500] : text;
    }
}

internal static class MacScreenshotHelperLocator
{
    private const uint MachO64Magic = 0xFEEDFACF;
    private const uint CpuTypeX64 = 0x01000007;
    private const uint CpuTypeArm64 = 0x0100000C;

    public static string Resolve(string baseDirectory)
    {
        if (!OperatingSystem.IsMacOS())
        {
            throw new PlatformNotSupportedException("The ScreenCaptureKit helper requires macOS.");
        }
        ArgumentException.ThrowIfNullOrWhiteSpace(baseDirectory);
        var path = Path.Combine(baseDirectory, "helpers", "excelmcp-screencapture");
        if (!File.Exists(path))
        {
            throw new FileNotFoundException(
                $"The bundled macOS screenshot helper is missing at '{path}'.", path);
        }
        var mode = File.GetUnixFileMode(path);
        const UnixFileMode executeBits =
            UnixFileMode.UserExecute | UnixFileMode.GroupExecute | UnixFileMode.OtherExecute;
        if ((mode & executeBits) == 0)
        {
            throw new InvalidOperationException(
                $"The bundled macOS screenshot helper is not executable: '{path}'.");
        }

        Span<byte> header = stackalloc byte[8];
        using (var stream = File.OpenRead(path))
        {
            if (stream.Read(header) != header.Length
                || BinaryPrimitives.ReadUInt32LittleEndian(header) != MachO64Magic)
            {
                throw new InvalidOperationException(
                    $"The bundled macOS screenshot helper is not a thin 64-bit Mach-O executable: '{path}'.");
            }
        }
        var actualCpuType = BinaryPrimitives.ReadUInt32LittleEndian(header[4..]);
        var expectedCpuType = RuntimeInformation.ProcessArchitecture switch
        {
            Architecture.Arm64 => CpuTypeArm64,
            Architecture.X64 => CpuTypeX64,
            var architecture => throw new PlatformNotSupportedException(
                $"The screenshot helper does not support process architecture '{architecture}'.")
        };
        if (actualCpuType != expectedCpuType)
        {
            throw new InvalidOperationException(
                $"The bundled macOS screenshot helper architecture does not match the current apphost.");
        }
        return path;
    }
}

internal sealed class MacScreenshotException(string code, string message) : InvalidOperationException(message)
{
    public string Code { get; } = code;
}
