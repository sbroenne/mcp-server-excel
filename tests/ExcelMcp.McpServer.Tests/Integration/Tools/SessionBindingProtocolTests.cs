using System.Text.Json;
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
public sealed class SessionBindingProtocolTests(ITestOutputHelper output)
    : McpIntegrationTestBase(output, "SessionBindingProtocolClient")
{
    [Theory]
    [InlineData("range", "get-used-range", true)]
    [InlineData("workbook", "get-info", true)]
    [InlineData("calculation_mode", "get-mode", true)]
    [InlineData("powerquery", "list", true)]
    [InlineData("file", "close", false)]
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
    public async Task MissingSessionId_ReturnsActionablePublicParameterName(string toolName, string action)
    {
        var invalidArguments = new[]
        {
            Arguments(action),
            new Dictionary<string, object?> { ["action"] = action, ["session_id"] = null },
            new Dictionary<string, object?> { ["action"] = action, ["session_id"] = "" },
            new Dictionary<string, object?> { ["action"] = action, ["session_id"] = "   " },
            new Dictionary<string, object?> { ["action"] = action, ["sessionId"] = "synthetic-private-value" },
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
    }

    [Fact]
    public async Task FileList_DoesNotRequireSessionIdentity()
    {
        var json = await CallToolAsync("file", Arguments("list"), TimeSpan.FromSeconds(30));
        using var document = ParseJsonResult(json, "file.list");
        Assert.True(document.RootElement.GetProperty("success").GetBoolean());
        Assert.Empty(document.RootElement.GetProperty("sessions").EnumerateArray());
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

    private static Dictionary<string, object?> Arguments(string action) => new()
    {
        ["action"] = action
    };
}
