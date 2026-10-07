using System.Text.Json;
using ModelContextProtocol.Protocol;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "File")]
[Trait("RequiresExcel", "false")]
public sealed class FileLifecycleContractProtocolTests : McpIntegrationTestBase
{
    public FileLifecycleContractProtocolTests(ITestOutputHelper output)
        : base(output, "FileLifecycleContractProtocolClient")
    {
    }

    [Fact]
    public async Task ListTools_FileSchema_ExposesOnlyCanonicalLifecycleActions()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var fileTool = Assert.Single(tools, tool => tool.Name == "file");
        var fileReadTool = Assert.Single(tools, tool => tool.Name == "file_read");
        var actions = fileTool.JsonSchema
            .GetProperty("properties")
            .GetProperty("action")
            .GetProperty("enum")
            .EnumerateArray()
            .Select(value => value.GetString()!)
            .ToArray();
        var readActions = fileReadTool.JsonSchema
            .GetProperty("properties")
            .GetProperty("action")
            .GetProperty("enum")
            .EnumerateArray()
            .Select(value => value.GetString()!)
            .ToArray();

        Assert.Equal(["open", "close", "create"], actions);
        Assert.Equal(["list", "test"], readActions);
        Assert.DoesNotContain("close-workbook", actions);
    }

    [Fact]
    public async Task FileCloseWorkbook_IsRejectedInsteadOfReportingNoOpSuccess()
    {
        var result = await Client!.CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "close-workbook",
            ["file_path"] = @"C:\tmp\book.xlsx"
        }, cancellationToken: TestCancellationToken);

        Assert.True(result.IsError);
        var response = Assert.Single(result.Content.OfType<TextContentBlock>()).Text;
        Output.WriteLine(response);
        Assert.False(string.IsNullOrWhiteSpace(response));
    }

    [Fact]
    public async Task FileCreate_UnsupportedExtension_ReturnsOriginalPathAndDoesNotCreateXlsx()
    {
        var tempDirectory = CreateTempDirectory("FileCreateUnsupportedExtension");
        var unsupportedPath = Path.Join(tempDirectory, $"UnsupportedCreate_{Guid.NewGuid():N}.txt");
        var renamedPath = Path.ChangeExtension(unsupportedPath, ".xlsx");
        string? sessionId = null;

        try
        {
            var result = await CallToolAsync("file", new Dictionary<string, object?>
            {
                ["action"] = "create",
                ["file_path"] = unsupportedPath
            });

            Output.WriteLine($"Unsupported file create result: {result}");

            using var json = JsonDocument.Parse(result);
            var root = json.RootElement;
            if (root.GetProperty("success").GetBoolean()
                && root.TryGetProperty("workbook_session_id", out var sessionIdProperty))
            {
                sessionId = sessionIdProperty.GetString();
                TrackSession(sessionId);
            }

            Assert.False(root.GetProperty("success").GetBoolean());
            Assert.True(root.GetProperty("isError").GetBoolean());
            Assert.Equal(unsupportedPath, root.GetProperty("filePath").GetString());

            var errorMessage = root.GetProperty("errorMessage").GetString();
            Assert.NotNull(errorMessage);
            Assert.Contains("Invalid file extension '.txt'", errorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(".xlsx", errorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(".xlsm", errorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(renamedPath), $"MCP file.create must not create '{renamedPath}'.");
        }
        finally
        {
            if (!string.IsNullOrWhiteSpace(sessionId))
            {
                await CloseSessionAsync(sessionId, save: false);
            }

            if (File.Exists(renamedPath))
            {
                File.Delete(renamedPath);
            }
        }
    }

    [Fact]
    public async Task FileTest_UsesTheSharedServiceResultShape()
    {
        var path = Path.Join(Path.GetTempPath(), $"missing-{Guid.NewGuid():N}.xlsx");
        using var service = new Service.ExcelMcpService();
        var serviceResponse = await service.ProcessAsync(new Service.ServiceRequest
        {
            Command = "session.test",
            Args = JsonSerializer.Serialize(new { filePath = path, timeoutSeconds = 120 })
        });
        var toolResult = await CallToolAsync("file_read", new() { ["action"] = "test", ["file_path"] = path });

        Assert.True(serviceResponse.Success);
        Assert.NotNull(serviceResponse.Result);
        Assert.Equal(serviceResponse.Result, toolResult);
        using var json = JsonDocument.Parse(toolResult);
        Assert.False(json.RootElement.GetProperty("canOpen").GetBoolean());
        Assert.False(json.RootElement.GetProperty("willOpenReadOnly").GetBoolean());
        Assert.False(json.RootElement.GetProperty("requiresVisibleSession").GetBoolean());
    }

    [Fact]
    public async Task FileTest_RelativePath_UsesSharedValidationError()
    {
        const string path = @"relative\book.xlsx";
        using var service = new Service.ExcelMcpService();
        var serviceResponse = await service.ProcessAsync(new Service.ServiceRequest
        {
            Command = "session.test",
            Args = JsonSerializer.Serialize(new { filePath = path, timeoutSeconds = 120 })
        });
        var toolResult = await CallToolAsync("file_read", new() { ["action"] = "test", ["file_path"] = path });

        Assert.False(serviceResponse.Success);
        Assert.Contains("absolute Windows path", serviceResponse.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        using var json = JsonDocument.Parse(toolResult);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(serviceResponse.ErrorMessage, json.RootElement.GetProperty("errorMessage").GetString());
    }
}
