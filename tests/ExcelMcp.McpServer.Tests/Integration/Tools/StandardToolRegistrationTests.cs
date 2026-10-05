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

    [Fact]
    public async Task Discovery_ReadOnlyToolsExposeOnlyReadActionsAndReadOnlyHint()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var byName = tools.ToDictionary(tool => tool.Name, StringComparer.Ordinal);
        var readTools = tools.Where(tool => tool.Name.EndsWith("_read", StringComparison.Ordinal)).ToArray();
        string ActionNames(string toolName) => string.Join(",",
            byName[toolName].JsonSchema.GetProperty("properties").GetProperty("action")
                .GetProperty("enum").EnumerateArray()
                .Select(value => value.GetString()).Order(StringComparer.Ordinal));

        Assert.NotEmpty(readTools);
        Assert.All(readTools, tool => Assert.True(
            tool.ProtocolTool.Annotations?.ReadOnlyHint == true,
            $"{tool.Name} must advertise readOnlyHint=true."));
        Assert.Equal("close,create,open", ActionNames("file"));
        Assert.Equal("list,test", ActionNames("file_read"));
        Assert.Equal("list", ActionNames("worksheet_read"));
        Assert.Equal("get-settings", ActionNames("calculation_mode_read"));
        Assert.DoesNotContain(
            byName["analysis_read"].JsonSchema.GetProperty("properties").GetProperty("action")
                .GetProperty("enum").EnumerateArray().Select(value => value.GetString()),
            action => action == "show-scenario");
        Assert.True(byName["screenshot"].ProtocolTool.Annotations?.ReadOnlyHint == true);
        Assert.DoesNotContain(
            byName["range"].JsonSchema.GetProperty("properties").GetProperty("action")
                .GetProperty("enum").EnumerateArray().Select(value => value.GetString()),
            action => action == "get-values");
        Assert.Contains(
            byName["range_read"].JsonSchema.GetProperty("properties").GetProperty("action")
                .GetProperty("enum").EnumerateArray().Select(value => value.GetString()),
            action => action == "get-values");
        Assert.Contains("evaluate", ActionNames("datamodel_read"));
        Assert.Contains("execute-dmv", ActionNames("datamodel_read"));
        Assert.Contains("find", ActionNames("range_edit_read"));
        Assert.Contains("preflight", ActionNames("table_read"));
    }

    [Fact]
    public async Task Discovery_WriteToolDescriptionsDoNotAdvertiseMovedReadActions()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var byName = tools.ToDictionary(tool => tool.Name, StringComparer.Ordinal);

        foreach (var readTool in tools.Where(tool => tool.Name.EndsWith("_read", StringComparison.Ordinal)))
        {
            var writeToolName = readTool.Name[..^"_read".Length];
            if (!byName.TryGetValue(writeToolName, out var writeTool))
                continue;

            Assert.DoesNotContain(readTool.Name, writeTool.Description, StringComparison.Ordinal);
            foreach (var action in readTool.JsonSchema.GetProperty("properties")
                .GetProperty("action").GetProperty("enum").EnumerateArray()
                .Select(value => value.GetString()!)
                .Where(action => action.Contains('-')))
            {
                Assert.DoesNotMatch(
                    $@"(?i)(?<![\w-]){Regex.Escape(action)}(?![\w-])",
                    writeTool.Description!);
            }
        }
    }

    [Theory]
    [InlineData("range_format", "validate-range", false)]
    [InlineData("connection", "test", true)]
    [InlineData("range", "trace-precedents", true)]
    [InlineData("range", "trace-dependents", true)]
    [InlineData("window", "get-view", false)]
    [InlineData("pythoninexcel", "get-result", false)]
    public async Task Discovery_InspectionActionsUseAccurateEndpoints(string toolName, string action, bool readOnly)
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var target = tools.Single(tool => tool.Name == (readOnly ? $"{toolName}_read" : toolName));
        var other = tools.SingleOrDefault(tool => tool.Name == (readOnly ? toolName : $"{toolName}_read"));
        static IEnumerable<string?> Actions(JsonElement schema) => schema.GetProperty("properties")
            .GetProperty("action").GetProperty("enum").EnumerateArray().Select(value => value.GetString());

        Assert.Contains(action, Actions(target.JsonSchema));
        if (other is not null)
            Assert.DoesNotContain(action, Actions(other.JsonSchema));
        Assert.Equal(readOnly, target.ProtocolTool.Annotations?.ReadOnlyHint == true);
    }

    [Fact]
    public async Task Discovery_InputSchemasUseJsonSchema202012OrItsProtocolDefault()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        const string dialect = "https://json-schema.org/draft/2020-12/schema";

        foreach (var tool in tools)
        {
            if (tool.JsonSchema.TryGetProperty("$schema", out var declaredDialect))
            {
                Assert.Equal(dialect, declaredDialect.GetString());
            }
        }
    }

    [Fact]
    public void Discovery_ServerVersionMatchesPackageInformationalVersion()
    {
        Assert.Equal(
            Infrastructure.McpServerVersionChecker.GetCurrentVersion(),
            Client!.ServerInfo.Version);
    }

    [Theory]
    [InlineData("range", "values")]
    [InlineData("range_format", "format_options")]
    [InlineData("workbook", "style_name")]
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
    public async Task LabelPositionDiscovery_AdvertisesOnlyWritablePositions()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var description = tools.Single(t => t.Name == "chart_config").JsonSchema
            .GetProperty("properties").GetProperty("label_position").GetProperty("description").GetString();
        Assert.NotNull(description);
        foreach (var name in new[] { "BestFit", "Center", "Above", "Below", "Left", "Right", "InsideBase", "InsideEnd", "OutsideEnd" })
        {
            Assert.Matches($@"\b{Regex.Escape(name)}\b", description);
        }
        Assert.DoesNotContain("Mixed", description, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("powerquery", "queryName", "query_name")]
    [InlineData("slicer", "slicerName", "slicer_name")]
    [InlineData("slicer", "destinationSheet", "destination_sheet")]
    [InlineData("slicer", "selectedItems", "selected_items")]
    [InlineData("slicer", "clearFirst", "clear_first")]
    public async Task RecoveryDescriptions_UseActualSdkInputNames(
        string toolName, string sourceName, string inputName)
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var tool = tools.Single(t => t.Name == toolName);
        Assert.True(tool.JsonSchema.GetProperty("properties").TryGetProperty(inputName, out _));
        Assert.Contains(inputName, tool.Description);
        Assert.DoesNotContain(sourceName, tool.Description);
    }

    [Fact]
    public async Task PercentageLabelDescription_DoesNotPromiseHarmlessUnsupportedWrites()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var description = tools.Single(t => t.Name == "chart_config").JsonSchema
            .GetProperty("properties").GetProperty("show_percentage").GetProperty("description").GetString();
        Assert.NotNull(description);
        Assert.DoesNotContain("no visual effect", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("pie", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("doughnut", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("reject", description, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task AxisSelectorDescription_DoesNotExcludeAcceptedLegacyAliases()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var description = tools.Single(t => t.Name == "chart_config").Description;
        Assert.NotNull(description);
        Assert.Contains("Primary=Category", description);
        Assert.Contains("Secondary=Value", description);
        Assert.Contains("both use the primary axis group", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("legacy", description, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task MeasureFormatDescription_ExplainsFailureInsteadOfGeneralSubstitution()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var description = tools.Single(t => t.Name == "datamodel").JsonSchema
            .GetProperty("properties").GetProperty("format_type").GetProperty("description").GetString();
        Assert.NotNull(description);
        Assert.Contains("not substituted", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("General", description);
        Assert.Contains("keeps the existing format", description, StringComparison.OrdinalIgnoreCase);
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
    [InlineData("range_read", "values")]
    [InlineData("range_read", "rowCount")]
    [InlineData("table_read", "tables")]
    [InlineData("worksheet", "worksheets")]
    [InlineData("file_read", "session_id")]
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
        var schema = Assert.IsType<JsonElement>(tools.Single(t => t.Name == "file_read").ReturnJsonSchema);
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
