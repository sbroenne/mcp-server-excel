using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Worksheets")]
[Trait("RequiresExcel", "true")]
public sealed class WorksheetImageToolTests : McpIntegrationTestBase
{
    public WorksheetImageToolTests(ITestOutputHelper output)
        : base(output, "WorksheetImageClient")
    {
    }

    [Fact]
    public async Task WorksheetStyle_AddImageAndCountImages_RoundsTripThroughMcp()
    {
        var tempDir = CreateTempDirectory("WorksheetImages");
        var workbookPath = Path.Combine(tempDir, "images.xlsx");
        var sessionId = await CreateWorkbookSessionAsync(workbookPath);
        await CreateWorksheetAsync(sessionId, "ImageSheet");

        var imagePath = Path.Combine(tempDir, "sample.png");
        File.WriteAllBytes(imagePath, Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAACklEQVR4nGMAAIAAeIhvAAAAAElFTkSuQmCC"));

        var addImageJson = await CallToolAsync("worksheet_style", new Dictionary<string, object?>
        {
            ["action"] = "add-image",
            ["session_id"] = sessionId,
            ["sheet_name"] = "ImageSheet",
            ["image_path"] = imagePath,
            ["cell_address"] = "A1"
        });
        AssertSuccess(addImageJson, "worksheet_style.add-image");

        var getImageCountJson = await CallToolAsync("worksheet_style", new Dictionary<string, object?>
        {
            ["action"] = "get-image-count",
            ["session_id"] = sessionId,
            ["sheet_name"] = "ImageSheet"
        });
        AssertSuccess(getImageCountJson, "worksheet_style.get-image-count");

        using var getImageCountDoc = JsonDocument.Parse(getImageCountJson);
        Assert.True(getImageCountDoc.RootElement.GetProperty("imageCount").GetInt32() > 0);

        await TryCloseSessionAsync(sessionId, save: true);
    }
}
