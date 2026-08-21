using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models.Actions;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Xunit;
using Xunit.Abstractions;
using ExcelServiceBridge = Sbroenne.ExcelMcp.McpServer.ServiceBridge.ServiceBridge;

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
        var actions = fileTool.JsonSchema
            .GetProperty("properties")
            .GetProperty("action")
            .GetProperty("enum")
            .EnumerateArray()
            .Select(value => value.GetString()!)
            .ToArray();

        Assert.Equal(["list", "open", "close", "create", "test"], actions);
        Assert.DoesNotContain("close-workbook", actions);
    }

    [Fact]
    public async Task FileCloseWorkbook_IsRejectedInsteadOfReportingNoOpSuccess()
    {
        var result = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "close-workbook",
            ["path"] = @"C:\tmp\book.xlsx"
        });

        Output.WriteLine(result);
        Assert.DoesNotContain("\"success\":true", result, StringComparison.OrdinalIgnoreCase);
        Assert.False(string.IsNullOrWhiteSpace(result));
    }

    [Fact]
    public async Task FileTest_UsesTheSharedServiceResultShape()
    {
        var path = Path.Join(Path.GetTempPath(), $"missing-{Guid.NewGuid():N}.xlsx");
        var serviceResponse = await ExcelServiceBridge.TestFileAsync(path);
        var toolResult = ExcelFileTool.ExcelFile(
            FileAction.Test,
            path,
            session_id: null,
            save: false,
            show: false,
            timeout_seconds: 120);

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
        var serviceResponse = await ExcelServiceBridge.TestFileAsync(path);
        var toolResult = ExcelFileTool.ExcelFile(
            FileAction.Test,
            path,
            session_id: null,
            save: false,
            show: false,
            timeout_seconds: 120);

        Assert.False(serviceResponse.Success);
        Assert.Contains("absolute Windows path", serviceResponse.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        using var json = JsonDocument.Parse(toolResult);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(serviceResponse.ErrorMessage, json.RootElement.GetProperty("errorMessage").GetString());
    }
}
