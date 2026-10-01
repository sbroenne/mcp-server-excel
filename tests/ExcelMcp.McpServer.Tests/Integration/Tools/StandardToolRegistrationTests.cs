using System.Text.Json;
using System.Text.RegularExpressions;
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
        Assert.Null(Client.ServerCapabilities.Prompts);
        Assert.Null(Client.ServerCapabilities.Resources);
    }

    [Theory]
    [InlineData("range", "values", "2D array")]
    [InlineData("range_format", "font_size", "points")]
    [InlineData("chart", "target_range", "left/top")]
    [InlineData("screenshot", "range_address", "capture")]
    [InlineData("screenshot", "quality", "JPEG")]
    public async Task ParameterDescriptions_PreserveCoreDocumentation(string toolName, string parameter, string detail)
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var description = tools.Single(t => t.Name == toolName).JsonSchema
            .GetProperty("properties").GetProperty(parameter).GetProperty("description").GetString();

        Assert.Contains(detail, description, StringComparison.OrdinalIgnoreCase);
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
    public async Task CalculationGuidance_RestoresPriorModeWithoutForcingUnrequestedChanges()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var description = tools.Single(t => t.Name == "calculation_mode").Description;

        Assert.Contains("restore the prior mode", description, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("set-mode(automatic)", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("restore the prior mode", Client.ServerInstructions, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("not mandatory", Client.ServerInstructions, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("does not request confirmation", Client.ServerInstructions, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("manual needs explicit calculate", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("what-if data tables, not worksheet Tables", description, StringComparison.Ordinal);
        Assert.Contains("asynchronous refreshes or Python calculations", description, StringComparison.Ordinal);
        Assert.Contains("concurrent requests and responses have no guaranteed order", Client.ServerInstructions, StringComparison.Ordinal);
        Assert.Contains("Await each dependent call", Client.ServerInstructions, StringComparison.Ordinal);
    }

    [Fact]
    public async Task FileDescription_ContainsCompactAuthorizedWorkflowAndDiscardWarning()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var tool = tools.Single(t => t.Name == "file");
        var description = tool.Description;
        Assert.Contains("Workflow: list and match", description, StringComparison.Ordinal);
        Assert.Contains("reuse its session or open/create", description, StringComparison.Ordinal);
        Assert.Contains("canClose", description, StringComparison.Ordinal);
        Assert.Contains("close when authorized with explicit save:true or save:false", description, StringComparison.Ordinal);
        var actionDescription = tool.JsonSchema.GetProperty("properties")
            .GetProperty("action").GetProperty("description").GetString();
        Assert.Contains("discards all unsaved edits, including earlier work", actionDescription, StringComparison.Ordinal);
        Assert.Contains("no tool-level undo", actionDescription, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("range", "clear-all removes values, formulas, and formats")]
    [InlineData("range", "clear-contents removes values/formulas")]
    [InlineData("range", "clear-formats removes formats")]
    [InlineData("worksheet", "removes all sheet contents and may break dependent references")]
    [InlineData("worksheet", "removes the source sheet and saves both files")]
    [InlineData("workbook", "replaces linked formulas with their current values")]
    public async Task DestructiveGuidance_ExposesConsequencesInDiscovery(string toolName, string consequence)
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var description = tools.Single(t => t.Name == toolName).Description;
        Assert.Contains(consequence, description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("no tool-level undo", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("retain", description, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Descriptions_UseAdvertisedTopLevelParameterNamesAndNoEmoji()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        foreach (var tool in tools)
        {
            var descriptions = EnumerateDescriptions(tool.JsonSchema).Prepend(tool.Description).ToArray();
            foreach (var description in descriptions)
            {
                Assert.False(description.EnumerateRunes().Any(rune =>
                    rune.Value is >= 0x1F000 and <= 0x1FAFF or >= 0x2600 and <= 0x27BF or 0xFE0F),
                    $"{tool.Name} description contains an emoji: {description}");
            }

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

    private static IEnumerable<string> EnumerateDescriptions(JsonElement element)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in element.EnumerateObject())
            {
                if (property.Name == "description" && property.Value.ValueKind == JsonValueKind.String)
                    yield return property.Value.GetString()!;
                else
                    foreach (var description in EnumerateDescriptions(property.Value))
                        yield return description;
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in element.EnumerateArray())
                foreach (var description in EnumerateDescriptions(item))
                    yield return description;
        }
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
