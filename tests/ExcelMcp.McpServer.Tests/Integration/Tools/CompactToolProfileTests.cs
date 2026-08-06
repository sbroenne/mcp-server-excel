using System.Text.Json;
using Sbroenne.ExcelMcp.Service.Workflow;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "ToolProfile")]
[Trait("Speed", "Fast")]
public sealed class CompactToolProfileTests : McpIntegrationTestBase
{
    private static readonly string[] ExpectedTools =
    [
        "calculation_mode",
        "file",
        "layout",
        "range",
        "range_edit",
        "range_format",
        "workflow",
        "worksheet",
        "worksheet_style",
    ];

    public CompactToolProfileTests(ITestOutputHelper output)
        : base(output, "CompactToolProfileClient")
    {
    }

    protected override IReadOnlyList<string> ServerArguments =>
        ["--tool-profile", "copilot-compact"];

    [Fact]
    public async Task ToolsList_ContainsOnlyTheStableCompactSurface()
    {
        Assert.NotNull(Client);
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var names = tools.Select(tool => tool.Name).Order(StringComparer.Ordinal).ToArray();

        Assert.Equal(ExpectedTools, names);
        Assert.All(tools, tool => Assert.Equal(JsonValueKind.Object, tool.JsonSchema.ValueKind));
        Assert.Equal(names.Length, names.Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public async Task Capabilities_ReportTheActualCompactProfile()
    {
        var result = await CallToolAsync("workflow", new Dictionary<string, object?>
        {
            ["action"] = "capabilities",
        });

        using var document = JsonDocument.Parse(result);
        var root = document.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean(), result);
        Assert.Equal("copilot-compact", root.GetProperty("toolProfile").GetString());
        Assert.Equal("1", root.GetProperty("toolProfileVersion").GetString());
        Assert.Equal("full", root.GetProperty("toolProfileFallback").GetString());
        Assert.Matches("^[0-9a-f]{64}$", root.GetProperty("toolProfileManifestHash").GetString());
        Assert.Equal(
            ExpectedTools.Order(StringComparer.Ordinal),
            root.GetProperty("toolProfileTools")
                .EnumerateArray()
                .Select(value => value.GetString()!)
                .Order(StringComparer.Ordinal));
        var expectedManifest = WorkflowRuntimeManifest.Create(
            typeof(Program).Assembly,
            "excel-mcp",
            McpToolProfileCatalog.CopilotCompactId,
            ExpectedTools,
            McpToolProfileCatalog.Version,
            McpToolProfileCatalog.FullId);
        Assert.Equal(expectedManifest.ToolProfileManifestHash, root.GetProperty("toolProfileManifestHash").GetString());
    }

    [Fact]
    public async Task FormattingSchemas_ExposeExplicitStateSettingAndReadbackActions()
    {
        Assert.NotNull(Client);
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);

        var layout = Assert.Single(tools, tool => tool.Name == "layout");
        var reportProperties = layout.JsonSchema.GetProperty("properties");
        Assert.Equal(
            ["apply-report", "get-report", "set-outline", "get-outline"],
            reportProperties.GetProperty("action").GetProperty("enum")
                .EnumerateArray().Select(value => value.GetString()!));
        Assert.Equal("string", reportProperties.GetProperty("header_range").GetProperty("type").GetString());
        Assert.Equal("string", reportProperties.GetProperty("body_range").GetProperty("type").GetString());
        Assert.Equal("boolean", reportProperties.GetProperty("auto_fit_columns").GetProperty("type").GetString());

        var outlineProperties = layout.JsonSchema.GetProperty("properties");
        Assert.Equal("string", outlineProperties.GetProperty("axis").GetProperty("type").GetString());
        Assert.Equal("integer", outlineProperties.GetProperty("level").GetProperty("type").GetString());
        Assert.Equal("boolean", outlineProperties.GetProperty("collapsed").GetProperty("type").GetString());

        Assert.False(reportProperties.TryGetProperty("review_only", out _));
        Assert.False(reportProperties.TryGetProperty("review_id", out _));
        Assert.False(reportProperties.TryGetProperty("checkpoint", out _));
        Assert.False(reportProperties.TryGetProperty("idempotency_key", out _));
    }
}
