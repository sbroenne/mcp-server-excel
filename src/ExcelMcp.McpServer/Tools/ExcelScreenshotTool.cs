using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Manual MCP tool for screenshot operations.
/// Returns ImageContentBlock for proper MCP image handling.
/// </summary>
[McpServerToolType]
public static class ExcelScreenshotTool
{
    /// <summary>
    /// Capture Excel worksheet content as images for visual verification.
    /// Uses Excel's built-in rendering to capture all visual elements (formatting, charts, conditional formatting).
    /// capture: specific range (requires rangeAddress).
    /// capture-sheet: entire used area of worksheet.
    /// Returns the image directly as MCP ImageContent.
    /// Use after operations to visually verify results.
    /// quality: Medium (default, JPEG 75% scale, ~4-8x smaller), High (PNG full scale), Low (JPEG 50% scale).
    /// </summary>
    [McpServerTool(Name = "screenshot", Title = "Screenshot", Destructive = false)]
    [McpMeta("category", "visualization")]
    [McpMeta("requiresSession", true)]
    [Description("Capture Excel worksheet content as images for visual verification. " +
        "Uses Excel's built-in rendering to capture all visual elements (formatting, charts, conditional formatting). " +
        "capture: specific range (requires rangeAddress). " +
        "capture-sheet: entire used area of worksheet. " +
        "Returns the image directly as MCP ImageContent. " +
        "Use after operations to visually verify results. " +
        "quality: Medium (default, JPEG 75% scale, ~4-8x smaller than High), High (PNG full scale), Low (JPEG 50% scale).")]
    public static CallToolResult ExcelScreenshot(
        [Description("The action to perform")] ScreenshotAction action,
        [Description("Session ID from file 'open' action")] string session_id,
        [DefaultValue(null)] string? sheet_name,
        [DefaultValue("A1:Z30")] string range_address,
        [DefaultValue(ScreenshotQuality.Medium)] ScreenshotQuality quality)
    {
        // Forward to service and get JSON response
        var jsonResponse = ExcelToolsBase.ExecuteToolAction(
            "screenshot",
            ServiceRegistry.Screenshot.ToActionString(action),
            () => ServiceRegistry.Screenshot.RouteAction(
                action,
                session_id,
                ExcelToolsBase.ForwardToServiceFunc,
                sheetName: sheet_name,
                rangeAddress: range_address,
                quality: quality
            ));

        // Parse the JSON response to extract image data
        try
        {
            var result = JsonSerializer.Deserialize<ScreenshotResult>(jsonResponse, ExcelToolsBase.JsonOptions);

            if (result is null || !result.Success || string.IsNullOrEmpty(result.ImageBase64))
            {
                // Return error as text content
                return new CallToolResult
                {
                    IsError = true,
                    Content = [new TextContentBlock { Text = jsonResponse }]
                };
            }

            // Return image as ImageContentBlock + metadata as TextContentBlock
            var metadata = $"Screenshot: {result.RangeAddress} on '{result.SheetName}' ({result.Width}x{result.Height}px)";

            return new CallToolResult
            {
                Content =
                [
                    ImageContentBlock.FromBytes(Convert.FromBase64String(result.ImageBase64), result.MimeType),
                    new TextContentBlock
                    {
                        Text = metadata
                    }
                ]
            };
        }
        catch (JsonException)
        {
            // If JSON parsing fails, return the raw response as error
            return new CallToolResult
            {
                IsError = true,
                Content = [new TextContentBlock { Text = jsonResponse }]
            };
        }
    }
}
