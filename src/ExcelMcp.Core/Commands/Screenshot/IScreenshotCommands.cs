using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Screenshot;

/// <summary>
/// Result containing a screenshot image as base64-encoded PNG data.
/// </summary>
public class ScreenshotResult : OperationResult
{
    /// <summary>Base64-encoded PNG image data</summary>
    public string ImageBase64 { get; set; } = string.Empty;

    /// <summary>MIME type of the image (always image/png)</summary>
    public string MimeType { get; set; } = "image/png";

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
/// RETURNS: Base64-encoded PNG image data with dimensions metadata.
/// For MCP: returned as inline ImageContent. For CLI: saved to file.
/// </summary>
[ServiceCategory("screenshot", "Screenshot")]
public interface IScreenshotCommands
{
    /// <summary>
    /// Captures a specific range as a PNG image.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name (null for active sheet)</param>
    /// <param name="rangeAddress">Range to capture (e.g., "A1:F20")</param>
    /// <returns>Screenshot result with base64 PNG data</returns>
    [ServiceAction("capture")]
    ScreenshotResult CaptureRange(IExcelBatch batch, string? sheetName = null, string rangeAddress = "A1:Z30");

    /// <summary>
    /// Captures the entire used area of a worksheet as a PNG image.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name (null for active sheet)</param>
    /// <returns>Screenshot result with base64 PNG data</returns>
    [ServiceAction("capture-sheet")]
    ScreenshotResult CaptureSheet(IExcelBatch batch, string? sheetName = null);
}
