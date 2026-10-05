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
    /// Photographs the live Excel window, so the image shows exactly what Excel displays
    /// (formatting, charts, conditional formatting). Works on protected sheets and leaves the
    /// workbook and clipboard untouched, but requires an interactive desktop session.
    /// capture: specific range (range_address defaults to A1:Z30).
    /// capture-sheet: used cell area of worksheet plus embedded charts.
    /// Very large areas are truncated to the top-left portion, which the returned text reports.
    /// Returns the image directly as MCP ImageContent.
    /// Use after operations to visually verify results.
    /// quality: Medium (default, JPEG 75% scale, ~4-8x smaller), High (PNG full scale), Low (JPEG 50% scale).
    /// </summary>
    [McpServerTool(Name = "screenshot", Title = "Screenshot", Destructive = false, ReadOnly = true,
        UseStructuredContent = true, OutputSchemaType = typeof(ScreenshotToolOutputSchema))]
    [McpMeta("category", "visualization")]
    [McpMeta("requiresSession", true)]
    [Description("Capture Excel worksheet content as images for visual verification. " +
        "Photographs the live Excel window, so the image shows exactly what Excel displays " +
        "(formatting, charts, conditional formatting). Works on protected sheets and leaves the " +
        "workbook and clipboard untouched, but requires an interactive desktop session. " +
        "capture: specific range (range_address defaults to A1:Z30). " +
        "capture-sheet: used cell area of worksheet plus embedded charts. " +
        "Very large areas are truncated to the top-left portion, which the returned text reports. " +
        "Returns the image directly as MCP ImageContent. " +
        "Use after operations to visually verify results. " +
        "quality: Medium (default, JPEG 75% scale, ~4-8x smaller than High), High (PNG full scale), Low (JPEG 50% scale).")]
    public static Task<CallToolResult> ExcelScreenshot(
        [Description("The action to perform")] ScreenshotAction action,
        ServiceBridge.ServiceBridge bridge,
        [Description("Session ID from file 'open' action")] string workbook_session_id,
        [Description("Worksheet name; omit for the active sheet. Valid for capture and capture-sheet.")]
        [DefaultValue(null)] string? sheet_name,
        [Description("Range to capture; defaults to A1:Z30. Only valid for capture, not capture-sheet.")]
        [DefaultValue("A1:Z30")] string range_address,
        [Description("Image quality: Medium (default, JPEG 75% scale), High (PNG full scale), Low (JPEG 50% scale).")]
        [DefaultValue(ScreenshotQuality.Medium)] ScreenshotQuality quality,
        CancellationToken cancellationToken = default)
    {
        return ExcelToolsBase.ExecuteToolActionAsync(
            "screenshot",
            ServiceRegistry.Screenshot.ToActionString(action),
            () => RouteScreenshotAction(
                action,
                workbook_session_id,
                sheet_name,
                range_address,
                quality,
                (command, id, args) => ExcelToolsBase.ForwardToServiceAsync(bridge, command, id, args, cancellationToken)
            ), cancellationToken, CreateToolResult);
    }

    internal static CallToolResult CreateToolResult(string jsonResponse)
    {
        var result = JsonSerializer.Deserialize<ScreenshotResult>(jsonResponse, ExcelToolsBase.JsonOptions)
            ?? throw new InvalidOperationException("Screenshot operation returned no result.");
        if (!result.Success)
            return ExcelToolsBase.CreateToolResult(jsonResponse, isError: true);
        if (string.IsNullOrEmpty(result.ImageBase64))
            throw new InvalidOperationException("Screenshot operation returned no image.");

        var metadata = $"Screenshot: {result.RangeAddress} on '{result.SheetName}' ({result.Width}x{result.Height}px)";

        // Preserve warnings such as truncation alongside the image.
        if (!string.IsNullOrWhiteSpace(result.Message))
        {
            metadata += $". {result.Message}";
        }

        return new CallToolResult
        {
            Content =
            [
                ImageContentBlock.FromBytes(Convert.FromBase64String(result.ImageBase64), result.MimeType),
                new TextContentBlock { Text = metadata }
            ],
            StructuredContent = JsonSerializer.SerializeToElement(new ScreenshotToolOutputSchema
            {
                Success = true,
                Message = result.Message,
                SheetName = result.SheetName,
                RangeAddress = result.RangeAddress,
                Width = result.Width,
                Height = result.Height,
                MimeType = result.MimeType
            }, ExcelToolsBase.JsonOptions)
        };
    }

    internal static TResult RouteScreenshotAction<TResult>(
        ScreenshotAction action,
        string sessionId,
        string? sheetName,
        string rangeAddress,
        ScreenshotQuality quality,
        Func<string, string, object?, TResult> forwardToService)
    {
        return ServiceRegistry.Screenshot.RouteAction(
            action,
            sessionId,
            forwardToService,
            sheetName: sheetName,
            rangeAddress: action == ScreenshotAction.CaptureRange ? rangeAddress : null,
            quality: quality);
    }
}
