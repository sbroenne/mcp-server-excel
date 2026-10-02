using System.ComponentModel;
using System.Reflection;
using System.Text.Json;
using System.Text.RegularExpressions;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "McpProtocol")]
[Trait("RequiresExcel", "false")]
public sealed class StandardToolRegistrationTests(ITestOutputHelper output)
    : McpIntegrationTestBase(output, "StandardRegistrationClient")
{
    [Fact]
    public async Task Discovery_PreservesToolsWithoutGeneratedGuidePrompts()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        Assert.Equal(McpToolSurface.ToolCount, tools.Count);
        Assert.False(string.IsNullOrWhiteSpace(Client.ServerInstructions));
        Assert.All(tools, tool => Assert.False(string.IsNullOrWhiteSpace(tool.Description)));
        Assert.Null(Client.ServerCapabilities.Prompts);
        Assert.Null(Client.ServerCapabilities.Resources);
    }

    [Theory]
    [InlineData("range", "values")]
    [InlineData("range_format", "font_size")]
    [InlineData("chart", "target_range")]
    [InlineData("screenshot", "range_address")]
    [InlineData("screenshot", "quality")]
    public async Task ParameterDescriptions_PreserveDeclaredMetadata(string toolName, string parameter)
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var description = tools.Single(t => t.Name == toolName).JsonSchema
            .GetProperty("properties").GetProperty(parameter).GetProperty("description").GetString();

        var metadata = GeneratedToolContract.GetParameter(toolName, parameter)
            .GetCustomAttribute<DescriptionAttribute>();
        Assert.NotNull(metadata);
        Assert.Equal(metadata.Description, description);
    }

    [Theory]
    [InlineData("axis", typeof(ChartAxisType))]
    [InlineData("legend_position", typeof(LegendPosition))]
    [InlineData("marker_style", typeof(MarkerStyle))]
    [InlineData("label_position", typeof(DataLabelPosition))]
    [InlineData("plot_by", typeof(ChartPlotBy))]
    [InlineData("display_blanks_as", typeof(ChartDisplayBlanksAs))]
    [InlineData("area", typeof(ChartAreaTarget))]
    public async Task EnumParameterDiscovery_AdvertisesEveryAcceptedName(string parameter, Type enumType)
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var description = tools.Single(t => t.Name == "chart_config").JsonSchema
            .GetProperty("properties").GetProperty(parameter).GetProperty("description").GetString();

        Assert.NotNull(description);
        foreach (var name in Enum.GetNames(enumType))
        {
            Assert.Matches($@"\b{Regex.Escape(name)}\b", description);
        }
    }

    [Fact]
    public async Task EveryInput_HasSubstantiveDescription()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        foreach (var tool in tools)
        {
            foreach (var parameter in tool.JsonSchema.GetProperty("properties").EnumerateObject())
            {
                Assert.True(parameter.Value.TryGetProperty("description", out var description),
                    $"{tool.Name}.{parameter.Name} has no description.");
                var text = description.GetString();
                Assert.False(string.IsNullOrWhiteSpace(text)
                    || text.StartsWith("(required", StringComparison.Ordinal)
                    || text.StartsWith("(valid", StringComparison.Ordinal),
                    $"{tool.Name}.{parameter.Name} has no description beyond action applicability.");
            }
        }
    }

    [Fact]
    public async Task Descriptions_UseAdvertisedTopLevelParameterNames()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        foreach (var tool in tools)
        {
            foreach (var property in tool.JsonSchema.GetProperty("properties").EnumerateObject())
            {
                var camelCase = Regex.Replace(property.Name, "_([a-z])", match => match.Groups[1].Value.ToUpperInvariant());
                if (camelCase == property.Name || camelCase == "sessionId")
                    continue; // sessionId is a local variable name in authored workflow examples.

                // Nested JSON object keys keep their schema spelling; inspect only top-level prose.
                var prose = new[] { tool.Description }.Concat(
                    tool.JsonSchema.GetProperty("properties").EnumerateObject()
                        .Where(p => p.Value.TryGetProperty("description", out _))
                        .Select(p => p.Value.GetProperty("description").GetString()!));
                foreach (var description in prose)
                {
                    var topLevelProse = Regex.Replace(description, @"\{[^{}]*\}", string.Empty);
                    Assert.False(Regex.IsMatch(topLevelProse, $@"\b{Regex.Escape(camelCase)}\b"),
                        $"{tool.Name} describes '{camelCase}' instead of '{property.Name}': {description}");
                }
            }
        }
    }

    [Fact]
    public async Task ParameterNameRendering_PreservesNestedJsonKeys()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var tool = tools.Single(t => t.Name == "table_column");

        Assert.Contains("{columnName, ascending}", tool.Description);
        Assert.DoesNotContain("{column_name, ascending}", tool.Description);
        var nested = tool.JsonSchema.GetProperty("properties").GetProperty("sort_columns")
            .GetProperty("items").GetProperty("properties");
        Assert.True(nested.TryGetProperty("columnName", out _));
    }

    [Fact]
    public async Task InjectedParameters_AreNotAdvertisedAsInputs()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        foreach (var tool in tools)
        {
            var properties = tool.JsonSchema.GetProperty("properties");
            foreach (var name in new[] { "bridge", "cancellationToken", "progress", "requestContext" })
            {
                Assert.False(properties.TryGetProperty(name, out _), $"{tool.Name} advertises {name}.");
            }
        }
    }

    [Fact]
    public async Task EveryTool_AdvertisesStructuredOutputSchema()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);

        foreach (var tool in tools)
        {
            var schema = Assert.IsType<JsonElement>(tool.ReturnJsonSchema);
            var properties = schema.GetProperty("properties");
            Assert.True(properties.TryGetProperty("success", out _),
                $"{tool.Name} output schema does not describe success.");
            Assert.True(properties.TryGetProperty("session_id", out _),
                $"{tool.Name} output schema does not describe session error context.");
            Assert.False(properties.TryGetProperty("sessionId", out _));
        }
    }

    [Theory]
    [InlineData("range", "values")]
    [InlineData("range", "rowCount")]
    [InlineData("table", "tables")]
    [InlineData("worksheet", "worksheets")]
    [InlineData("file", "session_id")]
    [InlineData("screenshot", "mimeType")]
    public async Task OutputSchemas_DescribeActionSpecificFields(string toolName, string propertyName)
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var schema = Assert.IsType<JsonElement>(tools.Single(t => t.Name == toolName).ReturnJsonSchema);

        Assert.True(
            schema.GetProperty("properties").TryGetProperty(propertyName, out _),
            $"{toolName} output schema does not describe '{propertyName}'.");
    }

    [Fact]
    public async Task FileListOutputSchema_DescribesSessionEntries()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var schema = Assert.IsType<JsonElement>(tools.Single(t => t.Name == "file").ReturnJsonSchema);
        var sessions = schema.GetProperty("properties").GetProperty("sessions");

        var properties = sessions.GetProperty("items").GetProperty("properties");
        foreach (var name in new[] { "session_id", "filePath", "isExcelVisible", "activeOperations", "canClose" })
        {
            Assert.True(properties.TryGetProperty(name, out _),
                $"File session output schema does not describe {name}.");
        }
        Assert.False(properties.TryGetProperty("sessionId", out _));
    }

    [Fact]
    public async Task ScreenshotOutputSchema_DescribesFailureMessage()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var schema = Assert.IsType<JsonElement>(tools.Single(t => t.Name == "screenshot").ReturnJsonSchema);

        Assert.True(schema.GetProperty("properties").TryGetProperty("errorMessage", out _),
            "Screenshot output schema does not describe its failure message.");

        var result = await Client.CallToolAsync("screenshot", new Dictionary<string, object?>
        {
            ["action"] = "capture",
            ["session_id"] = "unknown-screenshot-schema-session"
        }, cancellationToken: TestCancellationToken);

        Assert.True(result.IsError);
        var content = Assert.IsType<JsonElement>(result.StructuredContent);
        Assert.False(content.GetProperty("success").GetBoolean());
        Assert.False(string.IsNullOrWhiteSpace(content.GetProperty("errorMessage").GetString()));
    }

}
