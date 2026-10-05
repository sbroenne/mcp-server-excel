using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "CalculationMode")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class CalculationSettingsProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task Settings_MapsOnlySuppliedFields()
    {
        var expected = JsonSerializer.Serialize(new
        {
            iterationEnabled = false,
            maximumIterations = 37,
            maximumChange = 0.0002,
            calculateBeforeSave = false
        }, ServiceProtocol.JsonOptions);
        var call = await fixture.CallToolAsync("calculation_mode", new()
        {
            ["action"] = "set-settings",
            ["session_id"] = "session-1",
            ["iteration_enabled"] = false,
            ["maximum_iterations"] = 37,
            ["maximum_change"] = 0.0002,
            ["calculate_before_save"] = false
        }, RecordingToolTest.Success("""{"success":true}"""), "calculation.set-settings", expected);
        Assert.False(call.Result.IsError);
    }

    [Theory]
    [InlineData("normal")]
    [InlineData("full")]
    [InlineData("rebuild")]
    public async Task Calculation_MapsExplicitNativeStrength(string kind)
    {
        var expected = JsonSerializer.Serialize(new { scope = "application", kind }, ServiceProtocol.JsonOptions);
        var call = await fixture.CallToolAsync("calculation_mode", new()
        {
            ["action"] = "calculate",
            ["session_id"] = "session-1",
            ["scope"] = "application",
            ["kind"] = kind
        }, RecordingToolTest.Success("""{"success":true}"""), "calculation.calculate", expected);
        Assert.False(call.Result.IsError);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Precision_MapsExplicitLossPermission(bool enabled)
    {
        var expected = JsonSerializer.Serialize(new
        {
            precisionAsDisplayed = enabled,
            allowPrecisionLoss = enabled
        }, ServiceProtocol.JsonOptions);
        var call = await fixture.CallToolAsync("calculation_mode", new()
        {
            ["action"] = "set-precision",
            ["session_id"] = "session-1",
            ["precision_as_displayed"] = enabled,
            ["allow_precision_loss"] = enabled
        }, RecordingToolTest.Success("""{"success":true}"""), "calculation.set-precision", expected);
        Assert.False(call.Result.IsError);
    }

    [Theory]
    [InlineData("sheet", null, "sheetName")]
    [InlineData("range", null, "sheetName")]
    [InlineData("range", "Sheet1", "rangeAddress")]
    public async Task Calculation_MissingTargetForwardsInvalidInputAsToolError(
        string scope, string? sheetName, string parameter)
    {
        Dictionary<string, object?> arguments = new()
        {
            ["action"] = "calculate",
            ["session_id"] = "session-1",
            ["scope"] = scope
        };
        if (sheetName is not null)
            arguments["sheet_name"] = sheetName;
        var expected = sheetName is null
            ? JsonSerializer.Serialize(new { scope }, ServiceProtocol.JsonOptions)
            : JsonSerializer.Serialize(new { scope, sheetName }, ServiceProtocol.JsonOptions);
        var call = await fixture.CallToolAsync("calculation_mode", arguments, new ServiceResponse
        {
            Success = false,
            Command = "calculation.calculate",
            SessionId = "session-1",
            ErrorCategory = "InvalidInput",
            ExceptionType = nameof(ArgumentException),
            ErrorMessage = $"{parameter} is required for calculation."
        }, "calculation.calculate", expected);

        Assert.True(call.Result.IsError);
        using var json = JsonDocument.Parse(call.JsonResult);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("InvalidInput", json.RootElement.GetProperty("errorCategory").GetString());
        Assert.Equal(nameof(ArgumentException), json.RootElement.GetProperty("exceptionType").GetString());
        Assert.Contains(parameter, json.RootElement.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("get-mode")]
    [InlineData("set-mode")]
    public async Task RemovedActions_DoNotDispatch(string action)
    {
        var result = await fixture.CallResultWithoutDispatchAsync("calculation_mode", new()
        {
            ["action"] = action,
            ["session_id"] = "session-1"
        });
        Assert.True(result.IsError);
    }

    [Fact]
    public async Task MissingPrecisionChoice_DoesNotDispatch()
    {
        var result = await fixture.CallResultWithoutDispatchAsync("calculation_mode", new()
        {
            ["action"] = "set-precision",
            ["session_id"] = "session-1"
        });
        Assert.True(result.IsError);
    }

    [Fact]
    public async Task Discovery_ExplainsApplicationScopeAndPermanentPrecisionLoss()
    {
        var tools = await fixture.ListToolsAsync();
        var tool = Assert.Single(tools, item => item.Name == "calculation_mode");
        Assert.Contains("ALL open workbooks", tool.Description, StringComparison.Ordinal);
        Assert.Contains("stored numeric precision is permanently lost", tool.Description, StringComparison.Ordinal);
        var read = Assert.Single(tools, item => item.Name == "calculation_mode_read");
        Assert.Contains("get-settings", read.Description, StringComparison.Ordinal);
        Assert.True(tool.JsonSchema.GetProperty("properties").TryGetProperty("maximum_change", out _));
        Assert.True(tool.JsonSchema.GetProperty("properties").TryGetProperty("allow_precision_loss", out _));
    }
}
