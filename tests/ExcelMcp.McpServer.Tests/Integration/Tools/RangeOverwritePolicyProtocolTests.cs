using System.ComponentModel;
using System.Reflection;
using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "Range")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class RangeOverwritePolicyProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Theory]
    [InlineData("set-values", null)]
    [InlineData("set-values", "reject-nonempty")]
    [InlineData("set-values", "allow")]
    [InlineData("set-formulas", null)]
    [InlineData("set-formulas", "reject-nonempty")]
    [InlineData("set-formulas", "allow")]
    [InlineData("copy", null)]
    [InlineData("copy", "reject-nonempty")]
    [InlineData("copy", "allow")]
    [InlineData("copy", null, "values")]
    [InlineData("copy", "reject-nonempty", "values")]
    [InlineData("copy", "allow", "values")]
    [InlineData("copy", null, "formulas")]
    [InlineData("copy", "reject-nonempty", "formulas")]
    [InlineData("copy", "allow", "formulas")]
    public async Task ContentAction_MapsOptionalPolicy(string action, string? policy, string pasteKind = "all")
    {
        var arguments = Arguments(action, pasteKind);
        var serviceArgs = ServiceArguments(action, pasteKind);
        if (policy is not null)
        {
            arguments["overwrite_policy"] = policy;
            serviceArgs["overwritePolicy"] = policy;
        }
        var call = await fixture.CallToolAsync("range", arguments,
            RecordingToolTest.Success("""{"success":true}"""),
            $"range.{action}", JsonSerializer.Serialize(serviceArgs, ServiceProtocol.JsonOptions));
        Assert.False(call.Result.IsError);
        using var json = JsonDocument.Parse(call.JsonResult);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
    }

    [Fact]
    public async Task Conflict_PreservesErrorAndMarksProtocolFailure()
    {
        const string message = "Cannot write to occupied cells on sheet 'Sheet1'. Conflicting addresses: $A$1.";
        var call = await fixture.CallToolAsync("range", Arguments("set-values"), new ServiceResponse
        {
            Success = false,
            ErrorCategory = "Conflict",
            ErrorMessage = message,
            ExceptionType = "OperationFailureException"
        }, "range.set-values", JsonSerializer.Serialize(ServiceArguments("set-values"), ServiceProtocol.JsonOptions));
        Assert.True(call.Result.IsError);
        using var json = JsonDocument.Parse(call.JsonResult);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("Conflict", json.RootElement.GetProperty("errorCategory").GetString());
        Assert.Equal(message, json.RootElement.GetProperty("errorMessage").GetString());
    }

    [Fact]
    public async Task Discovery_AdvertisesProtectedDefaultAndLimitations()
    {
        var tools = await fixture.ListToolsAsync();
        var range = Assert.Single(tools, tool => tool.Name == "range");
        Assert.True(range.JsonSchema.GetProperty("properties").TryGetProperty("overwrite_policy", out _));
        Assert.Contains("reject-nonempty", range.Description, StringComparison.Ordinal);
        Assert.Contains("Never automatically retry", range.Description, StringComparison.Ordinal);
        Assert.Contains("10 conflicting addresses", range.Description, StringComparison.Ordinal);
        var parameter = GeneratedToolContract.GetParameter("range", "overwrite_policy");
        Assert.True(parameter.IsOptional);
        Assert.Equal(typeof(string), parameter.ParameterType);
        Assert.Null(parameter.DefaultValue);
        Assert.Contains("reject-nonempty", parameter.GetCustomAttribute<DescriptionAttribute>()?.Description, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ReadAction_RejectsInapplicablePolicy()
    {
        var result = await fixture.CallResultWithoutDispatchAsync("range", new()
        {
            ["action"] = "get-values",
            ["session_id"] = "session-1",
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1",
            ["overwrite_policy"] = "allow"
        });
        Assert.True(result.IsError);
    }

    private static Dictionary<string, object?> Arguments(string action, string pasteKind = "all")
    {
        var args = new Dictionary<string, object?> { ["action"] = action, ["session_id"] = "session-1" };
        if (action.StartsWith("copy", StringComparison.Ordinal))
        {
            args["source_sheet"] = "Sheet1";
            args["source_range"] = "A1:B2";
            args["target_sheet"] = "Sheet1";
            args["target_range"] = "D1";
            args["paste_kind"] = pasteKind;
        }
        else
        {
            args["sheet_name"] = "Sheet1";
            args["range_address"] = "A1";
            if (action == "set-values")
                args["values"] = new List<List<int>> { new() { 1 } };
            else
                args["formulas"] = new List<List<string>> { new() { "=1" } };
        }
        return args;
    }

    private static Dictionary<string, object?> ServiceArguments(string action, string pasteKind = "all")
    {
        var args = new Dictionary<string, object?>();
        foreach (var (key, value) in Arguments(action, pasteKind))
        {
            if (key is "action" or "session_id")
                continue;
            var parts = key.Split('_');
            string serviceName = parts[0] + string.Concat(parts.Skip(1).Select(part => char.ToUpperInvariant(part[0]) + part[1..]));
            args[serviceName] = value;
        }
        return args;
    }
}
