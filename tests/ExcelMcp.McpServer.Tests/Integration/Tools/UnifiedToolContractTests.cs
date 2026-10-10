using System.ComponentModel;
using System.Reflection;
using System.Text.Json;
using System.Text.RegularExpressions;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "GeneratedContracts")]
[Trait("RequiresExcel", "false")]
public sealed class UnifiedToolContractTests(RecordingProgramTransportFixture fixture)
{
    private static readonly JsonSerializerOptions SchemaJsonOptions = new()
    {
        Converters = { new System.Text.Json.Serialization.JsonStringEnumConverter() }
    };

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RequiredFileOrValue_AcceptsEitherAlternative(bool useFile)
    {
        var arguments = new Dictionary<string, object?>
        {
            ["action"] = "evaluate",
            ["workbook_session_id"] = "synthetic-session",
            [useFile ? "m_code_file" : "m_code"] = "synthetic-contract-value"
        };
        var call = await fixture.CallToolAsync("powerquery", arguments,
            new ServiceResponse { Success = true, Result = """{"success":true}""" },
            "powerquery.evaluate",
            JsonSerializer.Serialize(new
            {
                mCode = useFile ? null : "synthetic-contract-value",
                mCodeFile = useFile ? "synthetic-contract-value" : null
            }, ServiceProtocol.JsonOptions));
        Assert.False(call.Result.IsError == true);
    }

    [Theory]
    [InlineData("omitted", null)]
    [InlineData("null", null)]
    [InlineData("default", false)]
    [InlineData("explicit", true)]
    public async Task OptionalBoolean_PreservesOmittedNullDefaultAndExplicitValues(string input, bool? value)
    {
        var arguments = new Dictionary<string, object?>
        {
            ["action"] = "set-precision",
            ["workbook_session_id"] = "synthetic-session",
            ["precision_as_displayed"] = false
        };
        if (input != "omitted")
            arguments["allow_precision_loss"] = value;
        var call = await fixture.CallToolAsync(
            "calculation_mode",
            arguments,
            new ServiceResponse { Success = true, Result = """{"success":true}""" },
            "calculationmode.set-precision",
            JsonSerializer.Serialize(new { precisionAsDisplayed = false, allowPrecisionLoss = value },
                ServiceProtocol.JsonOptions));
        Assert.False(call.Result.IsError == true);
    }

    [Fact]
    public async Task EveryDiscoveredAction_AdvertisesExactlyItsRuntimeParameters()
    {
        var tools = await fixture.ListToolsAsync();
        Assert.Equal(McpToolSurface.ToolCount, tools.Count);
        foreach (var tool in tools)
        {
            var properties = tool.JsonSchema.GetProperty("properties");
            var actionNames = properties.GetProperty("action").GetProperty("enum")
                .EnumerateArray().Select(value => value.GetString()!).ToArray();
            foreach (var action in actionNames)
                Assert.Equal(ExpectedParameters(tool.Name, action).Order(StringComparer.Ordinal),
                    McpActionContract.GetParameters(tool.Name, action).Select(parameter => parameter.Name)
                        .Order(StringComparer.Ordinal));
            var runtimeNames = actionNames.SelectMany(action =>
                    McpActionContract.GetParameters(tool.Name, action).Select(parameter => parameter.Name))
                .Distinct(StringComparer.Ordinal).Order(StringComparer.Ordinal);
            Assert.Equal(runtimeNames, properties.EnumerateObject().Select(property => property.Name)
                .Order(StringComparer.Ordinal));

            var method = typeof(Program).Assembly.GetTypes()
                .Where(type => type.GetCustomAttribute<McpServerToolTypeAttribute>() is not null)
                .SelectMany(type => type.GetMethods(BindingFlags.Public | BindingFlags.Static))
                .Single(method => method.GetCustomAttribute<McpServerToolAttribute>()?.Name == tool.Name);
            var methodParameters = method.GetParameters()
                .Where(parameter => properties.TryGetProperty(parameter.Name!, out _)).ToArray();
            Assert.Equal(properties.EnumerateObject().Count(), methodParameters.Length);
            foreach (var parameter in methodParameters)
            {
                var schema = properties.GetProperty(parameter.Name!);
                var expectedType = JsonType(parameter.ParameterType);
                Assert.Contains(expectedType, SchemaTypes(schema));
                var defaultAttribute = parameter.GetCustomAttribute<DefaultValueAttribute>();
                if (defaultAttribute is not null)
                    Assert.True(schema.TryGetProperty("default", out _),
                        $"{tool.Name}.{parameter.Name}: missing explicit schema default.");
                if (schema.TryGetProperty("default", out var advertisedDefault))
                {
                    var value = defaultAttribute is not null
                        ? defaultAttribute.Value
                        : parameter.HasDefaultValue ? parameter.DefaultValue : null;
                    var json = JsonSerializer.SerializeToElement(value, value?.GetType() ?? typeof(object), SchemaJsonOptions);
                    Assert.Equal(json.ToString(), advertisedDefault.ToString());
                }
            }
        }
    }

    [Fact]
    public async Task DiscoveredInputTypes_MatchCoreDeclarationsAcrossAllActions()
    {
        var tools = (await fixture.ListToolsAsync()).ToDictionary(tool => tool.Name, StringComparer.Ordinal);
        var checkedInputs = 0;
        foreach (var contract in typeof(ISheetCommands).Assembly.GetTypes()
            .Where(type => type.IsInterface && type.GetCustomAttribute<ServiceCategoryAttribute>() is not null))
        {
            var category = contract.GetCustomAttribute<ServiceCategoryAttribute>()!;
            var baseTool = contract.GetCustomAttribute<McpToolAttribute>()?.ToolName
                ?? (category.PascalName == "Sheet" ? "worksheet" : category.PascalName.ToLowerInvariant());
            foreach (var method in contract.GetMethods())
            {
                var action = method.GetCustomAttribute<ServiceActionAttribute>()?.Action
                    ?? Regex.Replace(method.Name, "(?<!^)[A-Z]", "-$0").ToLowerInvariant();
                var toolName = method.GetCustomAttribute<McpToolAttribute>()?.ToolName ?? baseTool;
                if (!tools.TryGetValue(toolName, out var tool) || !Actions(tool.JsonSchema).Contains(action))
                    tool = tools.GetValueOrDefault($"{toolName}_read");
                if (baseTool == "diag")
                    continue;
                Assert.NotNull(tool);
                Assert.Contains(action, Actions(tool.JsonSchema));
                var properties = tool.JsonSchema.GetProperty("properties");
                foreach (var parameter in method.GetParameters().Where(parameter =>
                    parameter.ParameterType.Name != "IExcelBatch" &&
                    !(parameter.ParameterType.IsGenericType &&
                      parameter.ParameterType.GetGenericTypeDefinition() == typeof(IProgress<>))))
                {
                    var name = parameter.GetCustomAttribute<FromStringAttribute>()?.ExposedName ?? parameter.Name!;
                    var type = Nullable.GetUnderlyingType(parameter.ParameterType) ?? parameter.ParameterType;
                    if (type == typeof(TimeSpan))
                        name += "Seconds";
                    var inputName = Regex.Replace(name, "(?<!^)[A-Z]", "_$0").ToLowerInvariant();
                    Assert.True(properties.TryGetProperty(inputName, out var schema), $"{tool.Name}.{action}: missing {inputName}");
                    var expectedType = type == typeof(TimeSpan) ? "integer" :
                        type == typeof(List<string>) || parameter.GetCustomAttribute<FileOrValueAttribute>() is not null
                            ? "string" : JsonType(type);
                    Assert.Contains(expectedType, SchemaTypes(schema));
                    if (parameter.GetCustomAttribute<FileOrValueAttribute>() is { } fileOrValue)
                    {
                        var runtimeInput = Assert.Single(McpActionContract.GetParameters(tool.Name, action),
                            input => input.Name == inputName);
                        var required = parameter.GetCustomAttribute<RequiredParameterAttribute>() is not null ||
                            !parameter.IsOptional;
                        Assert.Equal(required, runtimeInput.Required);
                        Assert.Equal(required
                            ? Regex.Replace($"{name}{fileOrValue.FileSuffix}", "(?<!^)[A-Z]", "_$0").ToLowerInvariant()
                            : null, runtimeInput.Alternative);
                    }
                    checkedInputs++;
                }
            }
        }
        Assert.True(checkedInputs > McpToolSurface.OperationCount);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task EveryRequiredActionInput_RejectsOmissionOrNullBeforeDispatch(bool explicitNull)
    {
        var tools = await fixture.ListToolsAsync();
        var checkedInputs = 0;
        foreach (var tool in tools)
        {
            var properties = tool.JsonSchema.GetProperty("properties");
            foreach (var action in Actions(tool.JsonSchema))
            {
                var contract = McpActionContract.GetParameters(tool.Name, action);
                foreach (var required in contract.Where(parameter => parameter.Required && parameter.Name != "action"))
                {
                    var arguments = contract.ToDictionary(
                        parameter => parameter.Name,
                        parameter => parameter.Name == "action" ? (object?)action : SampleValue(properties.GetProperty(parameter.Name)));
                    if (explicitNull)
                        arguments[required.Name] = null;
                    else
                        arguments.Remove(required.Name);
                    if (required.Alternative is not null)
                    {
                        if (explicitNull)
                            arguments[required.Alternative] = null;
                        else
                            arguments.Remove(required.Alternative);
                    }

                    var result = await fixture.CallResultWithoutDispatchAsync(tool.Name, arguments);
                    Assert.True(result.IsError, $"{tool.Name}.{action}.{required.Name} accepted missing input.");
                    var text = Assert.Single(result.Content.OfType<TextContentBlock>()).Text;
                    Assert.Contains(required.Name, text, StringComparison.Ordinal);
                    Assert.DoesNotContain("synthetic-contract-value", text, StringComparison.Ordinal);
                    checkedInputs++;
                }
            }
        }
        Assert.True(checkedInputs > McpToolSurface.OperationCount);
    }

    [Fact]
    public async Task EveryAction_RejectsInputsFromOtherActionsIncludingNullAndDefaults()
    {
        var tools = await fixture.ListToolsAsync();
        var checkedInputs = 0;
        foreach (var tool in tools)
        {
            var properties = tool.JsonSchema.GetProperty("properties");
            foreach (var action in Actions(tool.JsonSchema))
            {
                var contract = McpActionContract.GetParameters(tool.Name, action);
                var allowed = ExpectedParameters(tool.Name, action);
                foreach (var property in properties.EnumerateObject().Where(property => !allowed.Contains(property.Name)))
                {
                    foreach (var value in new[] { null, SampleValue(property.Value) })
                    {
                        var arguments = contract.ToDictionary(
                            parameter => parameter.Name,
                            parameter => parameter.Name == "action" ? (object?)action : SampleValue(properties.GetProperty(parameter.Name)));
                        arguments[property.Name] = value;
                        var result = await fixture.CallResultWithoutDispatchAsync(tool.Name, arguments);
                        Assert.True(result.IsError, $"{tool.Name}.{action} accepted {property.Name}.");
                        var text = Assert.Single(result.Content.OfType<TextContentBlock>()).Text;
                        Assert.Contains(property.Name, text, StringComparison.Ordinal);
                        Assert.DoesNotContain("synthetic-contract-value", text, StringComparison.Ordinal);
                        checkedInputs++;
                    }
                }
            }
        }
        Assert.True(checkedInputs > McpToolSurface.OperationCount);
    }

    private static IEnumerable<string> Actions(JsonElement schema) => schema.GetProperty("properties")
        .GetProperty("action").GetProperty("enum").EnumerateArray().Select(value => value.GetString()!);

    private static HashSet<string> ExpectedParameters(string tool, string action)
    {
        var names = new HashSet<string>(StringComparer.Ordinal) { "action" };
        if (tool is "file" or "file_read")
        {
            var method = typeof(Program).Assembly.GetTypes()
                .Where(type => type.GetCustomAttribute<McpServerToolTypeAttribute>() is not null)
                .SelectMany(type => type.GetMethods(BindingFlags.Public | BindingFlags.Static))
                .Single(method => method.GetCustomAttribute<McpServerToolAttribute>()?.Name == tool);
            names.UnionWith(method.GetParameters().Where(parameter =>
                parameter.GetCustomAttributes<McpActionParameterAttribute>().Any(attribute => attribute.Action == action))
                .Select(parameter => parameter.Name!));
            return names;
        }

        var source = Assert.Single(typeof(ISheetCommands).Assembly.GetTypes()
            .Where(type => type.IsInterface && type.GetCustomAttribute<ServiceCategoryAttribute>() is not null)
            .SelectMany(type => type.GetMethods().Select(method => (Type: type, Method: method))),
            source =>
            {
                var baseTool = source.Type.GetCustomAttribute<McpToolAttribute>()?.ToolName
                    ?? (source.Type == typeof(ISheetCommands) ? "worksheet" :
                        source.Type.GetCustomAttribute<ServiceCategoryAttribute>()!.PascalName.ToLowerInvariant());
                var sourceTool = source.Method.GetCustomAttribute<McpToolAttribute>()?.ToolName ?? baseTool;
                var sourceAction = source.Method.GetCustomAttribute<ServiceActionAttribute>()?.Action
                    ?? Regex.Replace(source.Method.Name, "(?<!^)[A-Z]", "-$0").ToLowerInvariant();
                return (tool == sourceTool || tool == $"{sourceTool}_read") && action == sourceAction;
            });
        foreach (var parameter in source.Method.GetParameters())
        {
            if (parameter.ParameterType.Name == "IExcelBatch")
            {
                if (source.Type.GetCustomAttribute<NoSessionAttribute>() is null)
                    names.Add("workbook_session_id");
                continue;
            }
            if (parameter.ParameterType.IsGenericType &&
                parameter.ParameterType.GetGenericTypeDefinition() == typeof(IProgress<>))
                continue;
            var fileOrValue = parameter.GetCustomAttribute<FileOrValueAttribute>();
            var name = fileOrValue is not null ? parameter.Name! :
                parameter.GetCustomAttribute<FromStringAttribute>()?.ExposedName ?? parameter.Name!;
            if ((Nullable.GetUnderlyingType(parameter.ParameterType) ?? parameter.ParameterType) == typeof(TimeSpan))
                name += "Seconds";
            names.Add(Regex.Replace(name, "(?<!^)[A-Z]", "_$0").ToLowerInvariant());
            if (fileOrValue is not null)
                names.Add(Regex.Replace($"{name}{fileOrValue.FileSuffix}", "(?<!^)[A-Z]", "_$0").ToLowerInvariant());
        }
        return names;
    }

    private static IEnumerable<string> SchemaTypes(JsonElement schema)
    {
        var type = schema.GetProperty("type");
        return type.ValueKind == JsonValueKind.Array
            ? type.EnumerateArray().Select(value => value.GetString()!)
            : [type.GetString()!];
    }

    private static string JsonType(Type type)
    {
        type = Nullable.GetUnderlyingType(type) ?? type;
        if (type == typeof(string) || type.IsEnum)
            return "string";
        if (type == typeof(bool))
            return "boolean";
        if (type == typeof(int) || type == typeof(long))
            return "integer";
        if (type == typeof(double) || type == typeof(float) || type == typeof(decimal))
            return "number";
        if (type.IsArray || (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(List<>)))
            return "array";
        return "object";
    }

    private static object? SampleValue(JsonElement schema)
    {
        if (schema.TryGetProperty("default", out var defaultValue) && defaultValue.ValueKind != JsonValueKind.Null)
            return defaultValue.Clone();
        if (schema.TryGetProperty("enum", out var values))
            return values.EnumerateArray().First(value => value.ValueKind != JsonValueKind.Null).Clone();
        return SchemaTypes(schema).First(type => type != "null") switch
        {
            "string" => "synthetic-contract-value",
            "boolean" => false,
            "integer" => 0,
            "number" => 0.0,
            "array" => Array.Empty<object>(),
            "object" => new Dictionary<string, object?>(),
            _ => throw new InvalidOperationException("Unsupported synthetic schema.")
        };
    }
}
