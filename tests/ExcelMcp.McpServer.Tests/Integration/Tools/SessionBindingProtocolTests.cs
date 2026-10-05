using System.IO.Pipelines;
using System.Text.Json;
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
    [InlineData("range_read", "get-used-range", true)]
    [InlineData("workbook_read", "get-info", true)]
    [InlineData("calculation_mode_read", "get-settings", true)]
    [InlineData("powerquery_read", "list", true)]
    [InlineData("file", "close", false)]
    [InlineData("worksheet_read", "list", true)]
    public async Task WorkbookSessionId_SchemaAndWireNameAgree(string toolName, string action, bool required)
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var tool = Assert.Single(tools, tool => tool.Name == toolName);
        var schema = tool.JsonSchema;
        var properties = schema.GetProperty("properties");
        Assert.True(properties.TryGetProperty("workbook_session_id", out _));
        Assert.False(properties.TryGetProperty("session_id", out _));
        Assert.Equal(required, schema.GetProperty("required").EnumerateArray()
            .Any(property => property.GetString() == "workbook_session_id"));
        var returnSchema = Assert.IsType<JsonElement>(tool.ReturnJsonSchema);
        Assert.True(returnSchema.GetProperty("properties").TryGetProperty("workbook_session_id", out _));

        // The dictionary is sent through tools/call, not a direct tool method invocation.
        var arguments = Arguments(action);
        if (toolName is "range" or "range_read")
        {
            arguments["sheet_name"] = "Sheet1";
        }
        arguments["workbook_session_id"] = "synthetic-unknown-session";
        var json = await CallToolAsync(toolName, arguments, TimeSpan.FromSeconds(30));
        using var document = ParseJsonResult(json, toolName);
        Assert.False(document.RootElement.GetProperty("success").GetBoolean());
        var error = document.RootElement.GetProperty("errorMessage").GetString();
        Assert.Contains("not found", error, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("required", error, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("range_read", "get-used-range")]
    [InlineData("file", "close")]
    [InlineData("worksheet_read", "list")]
    public async Task OldSessionIdInput_IsRejectedAsUnknownParameter(string toolName, string action)
    {
        var arguments = SessionArguments(toolName, action);
        arguments["session_id"] = "synthetic-old-session";

        var response = await Client!.CallToolAsync(toolName, arguments,
            cancellationToken: TestCancellationToken);
        var text = Assert.Single(response.Content.OfType<TextContentBlock>()).Text;
        using var document = ParseJsonResult(text, $"{toolName}.{action}");
        AssertFailureEnvelope(document.RootElement, $"{toolName}.{action}",
            nameof(ArgumentException), expectedErrorCategory: "InvalidInput");
        Assert.True(response.IsError);
        Assert.DoesNotContain("synthetic-old-session", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("file", "open")]
    [InlineData("file", "create")]
    [InlineData("file_read", "test")]
    public async Task FilePathActions_DoNotRequireSessionIdentity(string toolName, string action)
    {
        var json = await CallToolAsync(toolName, Arguments(action), TimeSpan.FromSeconds(30));
        using var document = ParseJsonResult(json, action);
        Assert.False(document.RootElement.GetProperty("success").GetBoolean());
        var error = document.RootElement.GetProperty("errorMessage").GetString();
        Assert.Contains("path", error, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("session", error, StringComparison.OrdinalIgnoreCase);

        var oldArguments = Arguments(action);
        oldArguments["session_id"] = "synthetic-old-session";
        json = await CallToolAsync(toolName, oldArguments, TimeSpan.FromSeconds(30));
        using var oldDocument = ParseJsonResult(json, action);
        Assert.False(oldDocument.RootElement.GetProperty("success").GetBoolean());
        error = oldDocument.RootElement.GetProperty("errorMessage").GetString();
        Assert.Contains("Unknown parameter 'session_id'", error, StringComparison.Ordinal);
    }

    [Fact]
    public async Task FileList_DoesNotRequireSessionIdentity()
    {
        var json = await CallToolAsync("file_read", Arguments("list"), TimeSpan.FromSeconds(30));
        using var document = ParseJsonResult(json, "file_read.list");
        Assert.True(document.RootElement.GetProperty("success").GetBoolean());
        Assert.Empty(document.RootElement.GetProperty("sessions").EnumerateArray());

        var oldArguments = Arguments("list");
        oldArguments["session_id"] = "synthetic-old-session";
        json = await CallToolAsync("file_read", oldArguments, TimeSpan.FromSeconds(30));
        using var oldDocument = ParseJsonResult(json, "file_read.list");
        Assert.False(oldDocument.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains("Unknown parameter 'session_id'",
            oldDocument.RootElement.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task RangeWithWorkbookSessionIdButMissingSheetName_ReachesActionValidation()
    {
        var arguments = Arguments("get-used-range");
        arguments["workbook_session_id"] = "synthetic-unknown-session";
        var json = await CallToolAsync("range_read", arguments, TimeSpan.FromSeconds(30));
        using var document = ParseJsonResult(json, "range_read.get-used-range");
        AssertFailureEnvelope(document.RootElement, "range_read.get-used-range",
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
    public async Task InvalidActionWithoutSessionId_ReturnsSafeActionGuidance(string toolName, string? action)
    {
        var arguments = action is null ? new Dictionary<string, object?>() : Arguments(action);
        var response = await Client!.CallToolAsync(toolName, arguments,
            cancellationToken: TestCancellationToken);
        Assert.True(response.IsError);
        var text = Assert.Single(response.Content.OfType<TextContentBlock>()).Text;
        using var document = ParseJsonResult(text, toolName);
        AssertFailureEnvelope(document.RootElement, toolName,
            nameof(ArgumentException), expectedErrorCategory: "InvalidInput");
        Assert.Contains("action", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("workbook")]
    [InlineData("file")]
    [InlineData("worksheet")]
    public async Task InvalidActionWithOldSessionId_ReturnsSafeActionGuidance(string toolName)
    {
        var arguments = Arguments("synthetic-unknown-action");
        arguments["session_id"] = "synthetic-private-value";
        var response = await Client!.CallToolAsync(toolName, arguments,
            cancellationToken: TestCancellationToken);
        Assert.True(response.IsError);
        var text = Assert.Single(response.Content.OfType<TextContentBlock>()).Text;
        using var document = ParseJsonResult(text, toolName);
        AssertFailureEnvelope(document.RootElement, toolName,
            nameof(ArgumentException), expectedErrorCategory: "InvalidInput");
        Assert.Contains("action", text, StringComparison.Ordinal);
        Assert.DoesNotContain("synthetic-private-value", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(42)]
    [InlineData(true)]
    public async Task NonStringRequiredWorkbookSessionId_ReturnsStructuredInputError(object value)
    {
        var arguments = Arguments("get-info");
        arguments["workbook_session_id"] = value;
        var response = await Client!.CallToolAsync("workbook_read", arguments,
            cancellationToken: TestCancellationToken);
        Assert.True(response.IsError);
        var text = Assert.Single(response.Content.OfType<TextContentBlock>()).Text;
        using var document = ParseJsonResult(text, "workbook_read.get-info");
        AssertFailureEnvelope(document.RootElement, "workbook_read.get-info",
            nameof(ArgumentException), expectedErrorCategory: "InvalidInput");
        Assert.Contains("workbook_session_id", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("42")]
    [InlineData("true")]
    [InlineData("{\"private\":\"synthetic-private-value\"}")]
    [InlineData("[\"synthetic-private-value\"]")]
    public async Task FileClose_NonStringWorkbookSessionId_ReturnsStructuredInputError(string valueJson)
    {
        var arguments = Arguments("close");
        arguments["workbook_session_id"] = JsonSerializer.Deserialize<JsonElement>(valueJson);
        var response = await Client!.CallToolAsync("file", arguments,
            cancellationToken: TestCancellationToken);
        Assert.True(response.IsError);
        var text = Assert.Single(response.Content.OfType<TextContentBlock>()).Text;
        using var document = ParseJsonResult(text, "file.close");
        AssertFailureEnvelope(document.RootElement, "file.close",
            nameof(ArgumentException), expectedErrorCategory: "InvalidInput");
        Assert.Contains("workbook_session_id", text, StringComparison.Ordinal);
        Assert.DoesNotContain("synthetic-private-value", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("open")]
    [InlineData("create")]
    [InlineData("list")]
    [InlineData("test")]
    public async Task FileOptionalIdentity_RejectsUnusedParameter(string action)
    {
        var arguments = Arguments(action);
        arguments["workbook_session_id"] = 42;
        var toolName = action is "list" or "test" ? "file_read" : "file";
        var response = await Client!.CallToolAsync(toolName, arguments,
            cancellationToken: TestCancellationToken);
        Assert.True(response.IsError);
        var text = Assert.Single(response.Content.OfType<TextContentBlock>()).Text;
        using var document = ParseJsonResult(text, $"file.{action}");
        AssertFailureEnvelope(document.RootElement, $"file.{action}",
            nameof(ArgumentException), expectedErrorCategory: "InvalidInput");
        Assert.Contains("workbook_session_id", text, StringComparison.Ordinal);
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
    public async Task WorksheetSessionActions_RejectInvalidWorkbookSessionId(string action)
    {
        var toolName = action == "list" ? "worksheet_read" : "worksheet";
        foreach (var valueJson in new[] { "42", "true", "{}", "[]", "null", "\"\"", "\"   \"" })
        {
            var arguments = Arguments(action);
            arguments["workbook_session_id"] = JsonSerializer.Deserialize<JsonElement>(valueJson);
            var response = await Client!.CallToolAsync(toolName, arguments,
                cancellationToken: TestCancellationToken);
            Assert.True(response.IsError);
            var text = Assert.Single(response.Content.OfType<TextContentBlock>()).Text;
            Assert.Contains("workbook_session_id", text, StringComparison.Ordinal);
        }

        var validArguments = Arguments(action);
        validArguments["workbook_session_id"] = "synthetic-unknown-session";
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
        var json = await CallToolAsync(toolName, validArguments, TimeSpan.FromSeconds(30));
        using var result = ParseJsonResult(json, $"{toolName}.{action}");
        Assert.False(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains("not found", result.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ListTools_OptionalSessionIdentityIsLimitedToKnownConditionalTools()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var optional = tools.Where(tool =>
            tool.JsonSchema.GetProperty("properties").TryGetProperty("workbook_session_id", out _) &&
            (!tool.JsonSchema.TryGetProperty("required", out var required) ||
             !required.EnumerateArray().Any(value => value.GetString() == "workbook_session_id")))
            .Select(tool => tool.Name).Order().ToArray();
        Assert.Equal(["file", "worksheet"], optional);
    }

    [Theory]
    [InlineData("copy-to-file")]
    [InlineData("move-to-file")]
    public async Task WorksheetFileActions_RejectUnusedIdentity(string action)
    {
        var json = await CallToolAsync("worksheet", Arguments(action), TimeSpan.FromSeconds(30));
        using var document = ParseJsonResult(json, $"worksheet.{action}");
        Assert.False(document.RootElement.GetProperty("success").GetBoolean());
        var error = document.RootElement.GetProperty("errorMessage").GetString();
        Assert.Contains("sourceFile", error, StringComparison.Ordinal);
        Assert.DoesNotContain("session", error, StringComparison.OrdinalIgnoreCase);

        var arguments = Arguments(action);
        arguments["workbook_session_id"] = 42;
        var response = await Client!.CallToolAsync("worksheet", arguments,
            cancellationToken: TestCancellationToken);
        Assert.True(response.IsError);
        Assert.Contains("workbook_session_id",
            Assert.Single(response.Content.OfType<TextContentBlock>()).Text, StringComparison.Ordinal);

        arguments = Arguments(action);
        arguments["workbook_session_id"] = "synthetic-session";
        response = await Client!.CallToolAsync("worksheet", arguments,
            cancellationToken: TestCancellationToken);
        var text = Assert.Single(response.Content.OfType<TextContentBlock>()).Text;
        using var sessionDocument = ParseJsonResult(text, $"worksheet.{action}");
        AssertFailureEnvelope(sessionDocument.RootElement, $"worksheet.{action}",
            nameof(ArgumentException), expectedErrorCategory: "InvalidInput");
        Assert.Contains("workbook_session_id", text, StringComparison.Ordinal);
        Assert.DoesNotContain("synthetic-session", text, StringComparison.Ordinal);
    }

    private static Dictionary<string, object?> SessionArguments(string toolName, string action)
    {
        var arguments = Arguments(action);
        if (toolName is "range" or "range_read")
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
