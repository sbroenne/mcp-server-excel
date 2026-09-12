using System.IO.Pipelines;
using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Client;
using ModelContextProtocol;
using ModelContextProtocol.Protocol;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "SessionBinding")]
[Trait("RequiresExcel", "false")]
public sealed class SessionBindingProtocolTests : IAsyncLifetime, IAsyncDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private McpClient? _client;
    private Task? _serverTask;
    private bool _disposed;

    public SessionBindingProtocolTests(ITestOutputHelper output)
    {
        _output = output;
    }

    private McpClient? Client => _client;

    private CancellationToken TestCancellationToken => _cts.Token;

    public async Task InitializeAsync()
    {
        (_client, _serverTask) = await ProgramTransportTestHost.StartAsync(
            _clientToServerPipe,
            _serverToClientPipe,
            _cts.Token,
            "SessionBindingProtocolClient");
    }

    public async Task DisposeAsync()
    {
        await DisposeAsyncCore();
    }

    async ValueTask IAsyncDisposable.DisposeAsync()
    {
        await DisposeAsyncCore();
        GC.SuppressFinalize(this);
    }

    private async Task DisposeAsyncCore()
    {
        if (_disposed)
        {
            return;
        }
        _disposed = true;

        await ProgramTransportTestHost.StopAsync(
            _client,
            _clientToServerPipe,
            _serverToClientPipe,
            _serverTask,
            _output,
            _cts);
        _cts.Dispose();
    }

    [Theory]
    [InlineData("range", "get-used-range", true)]
    [InlineData("workbook", "get-info", true)]
    [InlineData("calculation_mode", "get-mode", true)]
    [InlineData("powerquery", "list", true)]
    [InlineData("file", "close", false)]
    [InlineData("worksheet", "list", false)]
    public async Task SessionId_SchemaAndWireNameAgree(string toolName, string action, bool required)
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var schema = Assert.Single(tools, tool => tool.Name == toolName).JsonSchema;
        var properties = schema.GetProperty("properties");
        Assert.Equal("string", properties.GetProperty("session_id").GetProperty("type").GetString());
        Assert.False(properties.TryGetProperty("sessionId", out _));
        Assert.Equal(required, schema.GetProperty("required").EnumerateArray()
            .Any(property => property.GetString() == "session_id"));

        // The dictionary is sent through tools/call, not a direct tool method invocation.
        var arguments = Arguments(action);
        if (toolName == "range")
        {
            arguments["sheet_name"] = "Sheet1";
        }
        arguments["session_id"] = "synthetic-unknown-session";
        var json = await CallToolAsync(toolName, arguments, TimeSpan.FromSeconds(30));
        using var document = ParseJsonResult(json, toolName);
        Assert.False(document.RootElement.GetProperty("success").GetBoolean());
        var error = document.RootElement.GetProperty("errorMessage").GetString();
        Assert.Contains("not found", error, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("required", error, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("range", "get-used-range")]
    [InlineData("workbook", "get-info")]
    [InlineData("calculation_mode", "get-mode")]
    [InlineData("powerquery", "list")]
    [InlineData("file", "close")]
    [InlineData("worksheet", "list")]
    [InlineData("worksheet", "create")]
    [InlineData("worksheet", "rename")]
    [InlineData("worksheet", "delete")]
    [InlineData("worksheet", "move")]
    [InlineData("worksheet", "copy")]
    public async Task MissingSessionId_ReturnsActionablePublicParameterName(string toolName, string action)
    {
        var invalidArguments = new[]
        {
            Arguments(action),
            new Dictionary<string, object?> { ["action"] = action, ["session_id"] = null },
            new Dictionary<string, object?> { ["action"] = action, ["session_id"] = "" },
            new Dictionary<string, object?> { ["action"] = action, ["session_id"] = "   " },
            new Dictionary<string, object?>
            {
                ["action"] = action,
                ["parameters"] = new { session_id = "synthetic-private-value" }
            }
        };
        foreach (var arguments in invalidArguments)
        {
            var response = await Client!.CallToolAsync(toolName, arguments,
                cancellationToken: TestCancellationToken);
            var text = Assert.Single(response.Content.OfType<TextContentBlock>()).Text;
            using var document = ParseJsonResult(text, toolName);
            AssertFailureEnvelope(document.RootElement, toolName,
                nameof(ArgumentException), expectedErrorCategory: "InvalidInput");
            Assert.True(response.IsError);

            Assert.Contains("session_id", text, StringComparison.Ordinal);
            Assert.Contains("arguments", text, StringComparison.Ordinal);
            Assert.Contains("file", text, StringComparison.Ordinal);
            Assert.DoesNotContain("sessionId", text, StringComparison.Ordinal);
            Assert.DoesNotContain("synthetic-private-value", text, StringComparison.Ordinal);
        }
    }

    [Theory]
    [InlineData("range", "get-used-range")]
    [InlineData("workbook", "get-info")]
    [InlineData("calculation_mode", "get-mode")]
    [InlineData("powerquery", "list")]
    [InlineData("file", "close")]
    [InlineData("worksheet", "list")]
    [InlineData("worksheet", "create")]
    [InlineData("worksheet", "rename")]
    [InlineData("worksheet", "delete")]
    [InlineData("worksheet", "move")]
    [InlineData("worksheet", "copy")]
    public async Task CamelCaseSessionIdAlias_ReachesSessionLookup(string toolName, string action)
    {
        var arguments = SessionArguments(toolName, action);
        arguments["sessionId"] = "synthetic-unknown-session";

        var json = await CallToolAsync(toolName, arguments, TimeSpan.FromSeconds(30));
        using var document = ParseJsonResult(json, $"{toolName}.{action}");
        Assert.False(document.RootElement.GetProperty("success").GetBoolean());
        var error = document.RootElement.GetProperty("errorMessage").GetString();
        Assert.Contains("not found", error, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("required", error, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("workbook", "get-info")]
    [InlineData("file", "close")]
    [InlineData("worksheet", "list")]
    public async Task EqualCanonicalAndAliasSessionIds_ReachSessionLookup(string toolName, string action)
    {
        var arguments = SessionArguments(toolName, action);
        arguments["session_id"] = "synthetic-unknown-session";
        arguments["sessionId"] = "synthetic-unknown-session";

        var json = await CallToolAsync(toolName, arguments, TimeSpan.FromSeconds(30));
        using var document = ParseJsonResult(json, $"{toolName}.{action}");
        Assert.False(document.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains("not found", document.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("workbook", "get-info")]
    [InlineData("file", "close")]
    [InlineData("worksheet", "list")]
    public async Task ConflictingCanonicalAndAliasSessionIds_ReturnStructuredInputError(
        string toolName,
        string action)
    {
        var arguments = SessionArguments(toolName, action);
        arguments["session_id"] = "synthetic-canonical-private-value";
        arguments["sessionId"] = "synthetic-alias-private-value";

        var response = await Client!.CallToolAsync(toolName, arguments,
            cancellationToken: TestCancellationToken);
        var text = Assert.Single(response.Content.OfType<TextContentBlock>()).Text;
        using var document = ParseJsonResult(text, $"{toolName}.{action}");
        AssertFailureEnvelope(document.RootElement, $"{toolName}.{action}",
            nameof(ArgumentException), expectedErrorCategory: "InvalidInput");
        Assert.True(response.IsError);
        Assert.Contains("session_id", text, StringComparison.Ordinal);
        Assert.DoesNotContain("synthetic-canonical-private-value", text, StringComparison.Ordinal);
        Assert.DoesNotContain("synthetic-alias-private-value", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("null", "\"synthetic-alias-private-value\"")]
    [InlineData("\"\"", "\"synthetic-alias-private-value\"")]
    [InlineData("\"   \"", "\"synthetic-alias-private-value\"")]
    [InlineData("42", "\"synthetic-alias-private-value\"")]
    [InlineData("\"synthetic-canonical-private-value\"", "null")]
    [InlineData("\"synthetic-canonical-private-value\"", "\"\"")]
    [InlineData("\"synthetic-canonical-private-value\"", "\"   \"")]
    [InlineData("\"synthetic-canonical-private-value\"", "42")]
    public async Task MalformedCanonicalOrAliasWhenBothPresent_ReturnsStructuredInputError(
        string canonicalJson,
        string aliasJson)
    {
        var arguments = Arguments("get-info");
        arguments["session_id"] = JsonSerializer.Deserialize<JsonElement>(canonicalJson);
        arguments["sessionId"] = JsonSerializer.Deserialize<JsonElement>(aliasJson);

        var response = await Client!.CallToolAsync("workbook", arguments,
            cancellationToken: TestCancellationToken);
        var text = Assert.Single(response.Content.OfType<TextContentBlock>()).Text;
        using var document = ParseJsonResult(text, "workbook.get-info");
        AssertFailureEnvelope(document.RootElement, "workbook.get-info",
            nameof(ArgumentException), expectedErrorCategory: "InvalidInput");
        Assert.True(response.IsError);
        Assert.Contains("session_id", text, StringComparison.Ordinal);
        Assert.DoesNotContain("synthetic-canonical-private-value", text, StringComparison.Ordinal);
        Assert.DoesNotContain("synthetic-alias-private-value", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("null")]
    [InlineData("\"\"")]
    [InlineData("\"   \"")]
    [InlineData("42")]
    [InlineData("true")]
    [InlineData("{}")]
    [InlineData("[]")]
    public async Task MalformedAliasWithoutCanonical_ReturnsStructuredInputError(string aliasJson)
    {
        var arguments = Arguments("get-info");
        arguments["sessionId"] = JsonSerializer.Deserialize<JsonElement>(aliasJson);

        var response = await Client!.CallToolAsync("workbook", arguments,
            cancellationToken: TestCancellationToken);
        var text = Assert.Single(response.Content.OfType<TextContentBlock>()).Text;
        using var document = ParseJsonResult(text, "workbook.get-info");
        AssertFailureEnvelope(document.RootElement, "workbook.get-info",
            nameof(ArgumentException), expectedErrorCategory: "InvalidInput");
        Assert.True(response.IsError);
        Assert.Contains("session_id", text, StringComparison.Ordinal);
    }

    [Fact]
    public void AliasObservation_WritesPrivacySafeWarningOnlyToStandardError()
    {
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();
        var originalOut = Console.Out;
        var originalError = Console.Error;

        try
        {
            Console.SetOut(stdout);
            Console.SetError(stderr);

            var services = new ServiceCollection();
            services.AddLogging(Program.ConfigureStdioLogging);
            using var provider = services.BuildServiceProvider();
            var logger = provider
                .GetRequiredService<ILoggerFactory>()
                .CreateLogger("SessionIdentityFilterTest");

            SessionIdentityFilter.WriteAliasWarning(logger, "workbook", "get-info");
        }
        finally
        {
            Console.SetOut(originalOut);
            Console.SetError(originalError);
        }

        Assert.Empty(stdout.ToString());
        Assert.Contains("Compatibility sessionId alias observed", stderr.ToString(),
            StringComparison.Ordinal);
        Assert.Contains("workbook/get-info", stderr.ToString(), StringComparison.Ordinal);
        Assert.DoesNotContain("synthetic-private-value", stderr.ToString(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("open")]
    [InlineData("create")]
    [InlineData("test")]
    public async Task FilePathActions_DoNotRequireSessionIdentity(string action)
    {
        var json = await CallToolAsync("file", Arguments(action), TimeSpan.FromSeconds(30));
        using var document = ParseJsonResult(json, action);
        Assert.False(document.RootElement.GetProperty("success").GetBoolean());
        var error = document.RootElement.GetProperty("errorMessage").GetString();
        Assert.Contains("path", error, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("session", error, StringComparison.OrdinalIgnoreCase);

        var aliasArguments = Arguments(action);
        aliasArguments["sessionId"] = "synthetic-optional-alias";
        json = await CallToolAsync("file", aliasArguments, TimeSpan.FromSeconds(30));
        using var aliasDocument = ParseJsonResult(json, action);
        Assert.False(aliasDocument.RootElement.GetProperty("success").GetBoolean());
        error = aliasDocument.RootElement.GetProperty("errorMessage").GetString();
        Assert.Contains("path", error, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("session", error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task FileList_DoesNotRequireSessionIdentity()
    {
        var json = await CallToolAsync("file", Arguments("list"), TimeSpan.FromSeconds(30));
        using var document = ParseJsonResult(json, "file.list");
        Assert.True(document.RootElement.GetProperty("success").GetBoolean());
        Assert.Empty(document.RootElement.GetProperty("sessions").EnumerateArray());

        var aliasArguments = Arguments("list");
        aliasArguments["sessionId"] = "synthetic-optional-alias";
        json = await CallToolAsync("file", aliasArguments, TimeSpan.FromSeconds(30));
        using var aliasDocument = ParseJsonResult(json, "file.list");
        Assert.True(aliasDocument.RootElement.GetProperty("success").GetBoolean());
        Assert.Empty(aliasDocument.RootElement.GetProperty("sessions").EnumerateArray());
    }

    [Fact]
    public async Task RangeWithSessionIdButMissingSheetName_ReachesActionValidation()
    {
        var arguments = Arguments("get-used-range");
        arguments["session_id"] = "synthetic-unknown-session";
        var json = await CallToolAsync("range", arguments, TimeSpan.FromSeconds(30));
        using var document = ParseJsonResult(json, "range.get-used-range");
        AssertFailureEnvelope(document.RootElement, "range.get-used-range",
            nameof(ArgumentException), expectedErrorCategory: "InvalidInput");
        var error = document.RootElement.GetProperty("errorMessage").GetString();
        Assert.Contains("sheetName", error, StringComparison.Ordinal);
        Assert.DoesNotContain("session", error, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("workbook", null)]
    [InlineData("workbook", "synthetic-unknown-action")]
    [InlineData("file", null)]
    [InlineData("file", "synthetic-unknown-action")]
    [InlineData("worksheet", null)]
    [InlineData("worksheet", "synthetic-unknown-action")]
    public async Task InvalidActionWithoutSessionId_PreservesSdkError(string toolName, string? action)
    {
        var arguments = action is null ? new Dictionary<string, object?>() : Arguments(action);
        var response = await Client!.CallToolAsync(toolName, arguments,
            cancellationToken: TestCancellationToken);
        Assert.True(response.IsError);
        var text = Assert.Single(response.Content.OfType<TextContentBlock>()).Text;
        Assert.Equal($"An error occurred invoking '{toolName}'.", text);
    }

    [Theory]
    [InlineData("workbook")]
    [InlineData("file")]
    [InlineData("worksheet")]
    public async Task InvalidActionWithAlias_PreservesSdkError(string toolName)
    {
        var arguments = Arguments("synthetic-unknown-action");
        arguments["sessionId"] = "synthetic-private-value";
        var response = await Client!.CallToolAsync(toolName, arguments,
            cancellationToken: TestCancellationToken);
        Assert.True(response.IsError);
        var text = Assert.Single(response.Content.OfType<TextContentBlock>()).Text;
        Assert.Equal($"An error occurred invoking '{toolName}'.", text);
        Assert.DoesNotContain("synthetic-private-value", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(42)]
    [InlineData(true)]
    public async Task NonStringRequiredSessionId_ReturnsStructuredInputError(object value)
    {
        var arguments = Arguments("get-info");
        arguments["session_id"] = value;
        var response = await Client!.CallToolAsync("workbook", arguments,
            cancellationToken: TestCancellationToken);
        Assert.True(response.IsError);
        var text = Assert.Single(response.Content.OfType<TextContentBlock>()).Text;
        using var document = ParseJsonResult(text, "workbook.get-info");
        AssertFailureEnvelope(document.RootElement, "workbook.get-info",
            nameof(ArgumentException), expectedErrorCategory: "InvalidInput");
        Assert.Contains("session_id", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("42")]
    [InlineData("true")]
    [InlineData("{\"private\":\"synthetic-private-value\"}")]
    [InlineData("[\"synthetic-private-value\"]")]
    public async Task FileClose_NonStringSessionId_ReturnsStructuredInputError(string valueJson)
    {
        var arguments = Arguments("close");
        arguments["session_id"] = JsonSerializer.Deserialize<JsonElement>(valueJson);
        var response = await Client!.CallToolAsync("file", arguments,
            cancellationToken: TestCancellationToken);
        Assert.True(response.IsError);
        var text = Assert.Single(response.Content.OfType<TextContentBlock>()).Text;
        using var document = ParseJsonResult(text, "file.close");
        AssertFailureEnvelope(document.RootElement, "file.close",
            nameof(ArgumentException), expectedErrorCategory: "InvalidInput");
        Assert.Contains("session_id", text, StringComparison.Ordinal);
        Assert.DoesNotContain("synthetic-private-value", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("open")]
    [InlineData("create")]
    [InlineData("list")]
    [InlineData("test")]
    public async Task FileOptionalIdentity_MalformedValuePreservesSdkError(string action)
    {
        var arguments = Arguments(action);
        arguments["session_id"] = 42;
        var response = await Client!.CallToolAsync("file", arguments,
            cancellationToken: TestCancellationToken);
        Assert.True(response.IsError);
        Assert.Equal("An error occurred invoking 'file'.",
            Assert.Single(response.Content.OfType<TextContentBlock>()).Text);
    }

    [Fact]
    public async Task UnknownTool_PreservesProtocolError()
    {
        var exception = await Assert.ThrowsAsync<McpProtocolException>(async () =>
            await Client!.CallToolAsync("synthetic-unknown-tool", Arguments("get-info"),
                cancellationToken: TestCancellationToken));
        Assert.Equal(McpErrorCode.InvalidParams, exception.ErrorCode);
        Assert.DoesNotContain("session_id", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("list")]
    [InlineData("create")]
    [InlineData("rename")]
    [InlineData("delete")]
    [InlineData("move")]
    [InlineData("copy")]
    public async Task WorksheetSessionActions_RejectInvalidIdentity(string action)
    {
        foreach (var valueJson in new[] { "42", "true", "{}", "[]", "null", "\"\"", "\"   \"" })
        {
            var arguments = Arguments(action);
            arguments["session_id"] = JsonSerializer.Deserialize<JsonElement>(valueJson);
            var response = await Client!.CallToolAsync("worksheet", arguments,
                cancellationToken: TestCancellationToken);
            Assert.True(response.IsError);
            var text = Assert.Single(response.Content.OfType<TextContentBlock>()).Text;
            using var document = ParseJsonResult(text, $"worksheet.{action}");
            AssertFailureEnvelope(document.RootElement, $"worksheet.{action}",
                nameof(ArgumentException), expectedErrorCategory: "InvalidInput");
            Assert.Contains("session_id", text, StringComparison.Ordinal);
        }

        var validArguments = Arguments(action);
        validArguments["session_id"] = "synthetic-unknown-session";
        if (action is "create" or "delete" or "move")
        {
            validArguments["sheet_name"] = "Sheet1";
        }
        if (action == "rename")
        {
            validArguments["old_name"] = "Sheet1";
            validArguments["new_name"] = "Renamed";
        }
        if (action == "copy")
        {
            validArguments["source_name"] = "Sheet1";
            validArguments["target_name"] = "Copied";
        }
        if (action == "move")
        {
            validArguments["before_sheet"] = "Sheet2";
        }
        var json = await CallToolAsync("worksheet", validArguments, TimeSpan.FromSeconds(30));
        using var result = ParseJsonResult(json, $"worksheet.{action}");
        Assert.False(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains("not found", result.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ListTools_OptionalSessionIdentityIsLimitedToKnownConditionalTools()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var optional = tools.Where(tool =>
            tool.JsonSchema.GetProperty("properties").TryGetProperty("session_id", out _) &&
            (!tool.JsonSchema.TryGetProperty("required", out var required) ||
             !required.EnumerateArray().Any(value => value.GetString() == "session_id")))
            .Select(tool => tool.Name).Order().ToArray();
        Assert.Equal(["file", "worksheet"], optional);
    }

    [Theory]
    [InlineData("copy-to-file")]
    [InlineData("move-to-file")]
    public async Task WorksheetFileActions_PreserveOptionalIdentity(string action)
    {
        var json = await CallToolAsync("worksheet", Arguments(action), TimeSpan.FromSeconds(30));
        using var document = ParseJsonResult(json, $"worksheet.{action}");
        Assert.False(document.RootElement.GetProperty("success").GetBoolean());
        var error = document.RootElement.GetProperty("errorMessage").GetString();
        Assert.Contains("sourceFile", error, StringComparison.Ordinal);
        Assert.DoesNotContain("session", error, StringComparison.OrdinalIgnoreCase);

        var arguments = Arguments(action);
        arguments["session_id"] = 42;
        var response = await Client!.CallToolAsync("worksheet", arguments,
            cancellationToken: TestCancellationToken);
        Assert.True(response.IsError);
        Assert.Equal("An error occurred invoking 'worksheet'.",
            Assert.Single(response.Content.OfType<TextContentBlock>()).Text);

        arguments = Arguments(action);
        arguments["sessionId"] = "synthetic-optional-alias";
        response = await Client!.CallToolAsync("worksheet", arguments,
            cancellationToken: TestCancellationToken);
        var text = Assert.Single(response.Content.OfType<TextContentBlock>()).Text;
        using var aliasDocument = ParseJsonResult(text, $"worksheet.{action}");
        AssertFailureEnvelope(aliasDocument.RootElement, $"worksheet.{action}",
            nameof(ArgumentException), expectedErrorCategory: "InvalidInput");
        Assert.Contains("sourceFile", text, StringComparison.Ordinal);
        Assert.DoesNotContain("session", text, StringComparison.OrdinalIgnoreCase);
    }

    private static Dictionary<string, object?> SessionArguments(string toolName, string action)
    {
        var arguments = Arguments(action);
        if (toolName == "range")
        {
            arguments["sheet_name"] = "Sheet1";
        }
        if (toolName == "worksheet")
        {
            if (action is "create" or "delete" or "move")
            {
                arguments["sheet_name"] = "Sheet1";
            }
            if (action == "rename")
            {
                arguments["old_name"] = "Sheet1";
                arguments["new_name"] = "Renamed";
            }
            if (action == "copy")
            {
                arguments["source_name"] = "Sheet1";
                arguments["target_name"] = "Copied";
            }
            if (action == "move")
            {
                arguments["before_sheet"] = "Sheet2";
            }
        }
        return arguments;
    }

    private static Dictionary<string, object?> Arguments(string action) => new()
    {
        ["action"] = action
    };

    private async Task<string> CallToolAsync(
        string toolName,
        Dictionary<string, object?> arguments,
        TimeSpan? timeout = null)
    {
        Assert.NotNull(Client);
        var callTask = Client!.CallToolAsync(
            toolName,
            arguments,
            cancellationToken: TestCancellationToken).AsTask();
        var result = timeout.HasValue
            ? await callTask.WaitAsync(timeout.Value, TestCancellationToken)
            : await callTask;
        return Assert.Single(result.Content.OfType<TextContentBlock>()).Text;
    }

    private static JsonDocument ParseJsonResult(string jsonResult, string operationName)
    {
        Assert.True(
            jsonResult.TrimStart().StartsWith('{'),
            $"{operationName} returned a non-JSON response. Response: {jsonResult}");
        return JsonDocument.Parse(jsonResult);
    }

    private static void AssertFailureEnvelope(
        JsonElement root,
        string operationName,
        string expectedExceptionType,
        string? expectedErrorCategory = null)
    {
        Assert.False(root.GetProperty("success").GetBoolean(), $"{operationName} unexpectedly succeeded.");
        Assert.True(root.GetProperty("isError").GetBoolean(), $"{operationName} should return isError=true.");
        Assert.Equal(expectedExceptionType, root.GetProperty("exceptionType").GetString());
        var error = root.GetProperty("error").GetString();
        var errorMessage = root.GetProperty("errorMessage").GetString();
        Assert.False(string.IsNullOrWhiteSpace(errorMessage), $"{operationName} should return errorMessage.");
        Assert.Equal(errorMessage, error);
        Assert.Equal(expectedErrorCategory, root.GetProperty("errorCategory").GetString());
    }
}
