using System.Text.Json;
using System.Text.Json.Nodes;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

internal static class ScreenshotImageResult
{
    internal static CallToolResult CreateToolResult(string jsonResponse)
    {
        var result = JsonSerializer.Deserialize<ScreenshotResult>(jsonResponse, ExcelToolsBase.JsonOptions)
            ?? throw new InvalidOperationException("Screenshot operation returned no result.");
        if (!result.Success)
            return ExcelToolsBase.CreateToolResult(jsonResponse, isError: true);
        if (string.IsNullOrEmpty(result.ImageBase64))
            throw new InvalidOperationException("Screenshot operation returned no image.");

        var metadata = $"Screenshot: {result.RangeAddress} on '{result.SheetName}' ({result.Width}x{result.Height}px)";
        if (!string.IsNullOrWhiteSpace(result.Message))
            metadata += $". {result.Message}";
        var structured = JsonNode.Parse(jsonResponse)!.AsObject();
        structured.Remove("imageBase64");

        return new CallToolResult
        {
            Content =
            [
                ImageContentBlock.FromBytes(Convert.FromBase64String(result.ImageBase64), result.MimeType),
                new TextContentBlock { Text = metadata }
            ],
            StructuredContent = JsonSerializer.SerializeToElement(structured, ExcelToolsBase.JsonOptions)
        };
    }
}
