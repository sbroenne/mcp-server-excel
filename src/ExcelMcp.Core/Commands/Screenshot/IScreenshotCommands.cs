using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Screenshot;

/// <summary>
/// Image quality level for screenshot capture.
/// Controls the output format and scale to balance visual fidelity against response size.
/// </summary>
public enum ScreenshotQuality
{
    /// <summary>JPEG at 75% scale. Recommended for most uses. ~4-8x smaller than High.</summary>
    Medium = 0,
    /// <summary>PNG at full scale. Maximum fidelity for detailed inspection.</summary>
    High = 1,
    /// <summary>JPEG at 50% scale. Smallest size, good for overview/layout verification.</summary>
    Low = 2
}

/// <summary>
/// Result containing a screenshot image as base64-encoded image data.
/// </summary>
public class ScreenshotResult : OperationResult
{
    /// <summary>Base64-encoded image data</summary>
    public string ImageBase64 { get; set; } = string.Empty;

    /// <summary>MIME type of the image (image/png or image/jpeg)</summary>
    public string MimeType { get; set; } = "image/jpeg";

    /// <summary>Image width in pixels</summary>
    public int Width { get; set; }

    /// <summary>Image height in pixels</summary>
    public int Height { get; set; }

    /// <summary>Sheet name that was captured</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SheetName { get; set; }

    /// <summary>Range address that was captured (e.g., "A1:F20")</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? RangeAddress { get; set; }
}

/// <summary>
/// Capture Excel worksheet content as images for visual verification.
/// Uses Excel's built-in rendering (CopyPicture) to capture ranges as PNG images.
/// Captures formatting, conditional formatting, charts, and all visual elements.
///
/// ACTIONS:
/// - capture: Capture a specific range as an image
/// - capture-sheet: Capture the entire used area of a worksheet
///
/// RETURNS: Base64-encoded image data with dimensions metadata.
/// For MCP: returned as native ImageContent (no file handling needed).
/// For CLI: use --output &lt;path&gt; to save the image directly to a PNG/JPEG file instead of returning base64 inline.
/// Quality defaults to Medium (JPEG 75% scale) which is 4-8x smaller than High (PNG).
/// Use High only when fine detail inspection is needed.
/// </summary>
[ServiceCategory("screenshot", "Screenshot")]
public interface IScreenshotCommands
{
    /// <summary>
    /// Captures a specific range as an image.
    /// For CLI: use --output &lt;path&gt; to save the image directly to a PNG/JPEG file.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name (null for active sheet)</param>
    /// <param name="rangeAddress">Range to capture (e.g., "A1:F20")</param>
    /// <param name="quality">Image quality: Medium (default, JPEG 75% scale), High (PNG full scale), Low (JPEG 50% scale)</param>
    /// <returns>Screenshot result with base64 image data</returns>
    [ServiceAction("capture")]
    ScreenshotResult CaptureRange(IExcelBatch batch, string? sheetName = null, string rangeAddress = "A1:Z30", ScreenshotQuality quality = ScreenshotQuality.Medium);

    /// <summary>
    /// Captures the entire used area of a worksheet as an image.
    /// For CLI: use --output &lt;path&gt; to save the image directly to a PNG/JPEG file.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name (null for active sheet)</param>
    /// <param name="quality">Image quality: Medium (default, JPEG 75% scale), High (PNG full scale), Low (JPEG 50% scale)</param>
    /// <returns>Screenshot result with base64 image data</returns>
    [ServiceAction("capture-sheet")]
    ScreenshotResult CaptureSheet(IExcelBatch batch, string? sheetName = null, ScreenshotQuality quality = ScreenshotQuality.Medium);
}
