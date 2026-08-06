using System.Text.Json;
using Sbroenne.ExcelMcp.Service.Workflow;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Workflow")]
[Trait("Speed", "Fast")]
public sealed class WorkflowToolContractTests : McpIntegrationTestBase
{
    private static readonly string[] ExpectedActions = ["capabilities", "open-and-describe", "execute-plan"];

    public WorkflowToolContractTests(ITestOutputHelper output)
        : base(output, "WorkflowToolContractClient")
    {
    }

    [Fact]
    public async Task Capabilities_IdentifiesTheInstalledWorkflowSurface_ViaMcpProtocol()
    {
        var result = await CallToolAsync("workflow", new Dictionary<string, object?>
        {
            ["action"] = "capabilities",
        });

        using var document = JsonDocument.Parse(result);
        var root = document.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean());
        Assert.True(root.GetProperty("executePlan").GetBoolean());
        Assert.True(root.GetProperty("openAndDescribe").GetBoolean());
        Assert.Equal("2", root.GetProperty("workflowInterfaceVersion").GetString());
        Assert.Equal("excel-mcp", root.GetProperty("runtimeHost").GetString());
        Assert.Equal("full", root.GetProperty("toolProfile").GetString());
        Assert.Equal(Client!.ServerInfo!.Version, root.GetProperty("serverVersion").GetString());
        Assert.False(string.IsNullOrWhiteSpace(root.GetProperty("buildFingerprint").GetString()));
        Assert.Matches("^[0-9a-f]{64}$", root.GetProperty("toolProfileManifestHash").GetString());
        var expectedManifest = WorkflowRuntimeManifest.Create(
            typeof(Program).Assembly,
            "excel-mcp",
            McpToolProfileCatalog.FullId,
            GeminiCompatibleToolRegistration.DiscoverActiveToolNames(typeof(Program).Assembly, McpToolProfile.Full),
            McpToolProfileCatalog.Version,
            McpToolProfileCatalog.FullId);
        Assert.Equal(expectedManifest.BuildFingerprint, root.GetProperty("buildFingerprint").GetString());
        Assert.Equal(expectedManifest.ToolProfileManifestHash, root.GetProperty("toolProfileManifestHash").GetString());
        Assert.Equal(expectedManifest.ToolProfileTools, root.GetProperty("toolProfileTools")
            .EnumerateArray().Select(value => value.GetString()));
        Assert.True(root.GetProperty("compactReceipts").GetBoolean());
        Assert.True(root.GetProperty("planCheckpoint").GetBoolean());
        Assert.True(root.GetProperty("planIdempotency").GetBoolean());
        Assert.True(root.GetProperty("finalRangeVerification").GetBoolean());
        Assert.False(root.GetProperty("planReview").GetBoolean());
        Assert.True(root.GetProperty("fastMode").GetBoolean());
        Assert.Equal("1", root.GetProperty("fastModeVersion").GetString());
        Assert.Equal("sequential", root.GetProperty("fastModeFallback").GetString());
        Assert.Contains("range", root.GetProperty("fastModeCompatibleCategories")
            .EnumerateArray().Select(value => value.GetString()));
    }

    [Fact]
    public async Task Schema_DescribesConstructibleOrderedOperations_ViaMcpProtocol()
    {
        Assert.NotNull(Client);
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var workflow = Assert.Single(tools, tool => tool.Name == "workflow");
        var schema = workflow.JsonSchema;
        var properties = schema.GetProperty("properties");

        var actionValues = properties.GetProperty("action").GetProperty("enum")
            .EnumerateArray()
            .Select(value => value.GetString()!)
            .ToArray();
        Assert.Equal(ExpectedActions, actionValues);

        var operations = properties.GetProperty("operations");
        Assert.Equal("array", operations.GetProperty("type").GetString());
        var operationItems = operations.GetProperty("items");
        Assert.Equal("object", operationItems.GetProperty("type").GetString());
        var operationProperties = operationItems.GetProperty("properties");
        Assert.Equal("string", operationProperties.GetProperty("command").GetProperty("type").GetString());
        Assert.Equal("object", operationProperties.GetProperty("args").GetProperty("type").GetString());
        var checkpointMode = properties.GetProperty("checkpoint_mode");
        Assert.Contains("inherit", checkpointMode.GetProperty("enum").EnumerateArray().Select(v => v.GetString()));
        Assert.Contains("off", checkpointMode.GetProperty("enum").EnumerateArray().Select(v => v.GetString()));
        Assert.Contains("once", checkpointMode.GetProperty("enum").EnumerateArray().Select(v => v.GetString()));
        Assert.Equal("string", properties.GetProperty("idempotency_key").GetProperty("type").GetString());
        Assert.Equal("boolean", properties.GetProperty("fast_mode").GetProperty("type").GetString());
        Assert.True(properties.GetProperty("fast_mode").GetProperty("default").GetBoolean());
        Assert.Equal("string", properties.GetProperty("verify_sheet_name").GetProperty("type").GetString());
        Assert.Equal("string", properties.GetProperty("verify_range_address").GetProperty("type").GetString());
        Assert.Contains(
            "rangeformat.format-range",
            operationProperties.GetProperty("command").GetProperty("description").GetString(),
            StringComparison.Ordinal);
        Assert.Contains(
            "command",
            operationItems.GetProperty("required").EnumerateArray().Select(value => value.GetString()!));
    }

    [Fact]
    public async Task ExecutePlan_RejectsAnIncompleteVerificationScope_BeforeDispatch()
    {
        var result = await CallToolAsync("workflow", new Dictionary<string, object?>
        {
            ["action"] = "execute-plan",
            ["session_id"] = "not-a-real-session",
            ["operations"] = new object?[]
            {
                new Dictionary<string, object?>
                {
                    ["command"] = "range.get-values",
                    ["args"] = new Dictionary<string, object?>
                    {
                        ["sheetName"] = "Data",
                        ["rangeAddress"] = "A1",
                    },
                },
            },
            ["verify_sheet_name"] = "Data",
        });

        using var document = JsonDocument.Parse(result);
        Assert.False(document.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains(
            "verify_sheet_name and verify_range_address must be supplied together",
            document.RootElement.GetProperty("error").GetString(),
            StringComparison.Ordinal);
    }
}
