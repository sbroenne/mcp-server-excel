using System.Text.Json;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Telemetry")]
public sealed class ExcelToolsBaseTelemetryTests
{
    [Fact]
    public void ExecuteToolAction_SuccessResponse_TracksSuccess()
    {
        ToolInvocation? invocation = null;

        var response = ExcelToolsBase.ExecuteToolAction(
            "range",
            "get-values",
            path: null,
            operation: () => """{"success":true,"value":"private workbook content"}""",
            trackInvocation: (tool, action, duration, success, path) =>
                invocation = new ToolInvocation(tool, action, duration, success, path));

        Assert.Equal("""{"success":true,"value":"private workbook content"}""", response);
        Assert.NotNull(invocation);
        Assert.True(invocation.Success);
    }

    [Fact]
    public void ExecuteToolAction_StructuredFailureResponse_TracksFailure()
    {
        ToolInvocation? invocation = null;

        var response = ExcelToolsBase.ExecuteToolAction(
            "range",
            "get-values",
            path: null,
            operation: () => """{"success":false,"errorMessage":"private workbook details"}""",
            trackInvocation: (tool, action, duration, success, path) =>
                invocation = new ToolInvocation(tool, action, duration, success, path));

        Assert.Equal("""{"success":false,"errorMessage":"private workbook details"}""", response);
        Assert.NotNull(invocation);
        Assert.False(invocation.Success);
    }

    [Fact]
    public void ExecuteToolAction_PrimitiveJsonResponse_TracksSuccess()
    {
        ToolInvocation? invocation = null;

        var response = ExcelToolsBase.ExecuteToolAction(
            "chart",
            "get-axis-number-format",
            path: null,
            operation: () => "\"General\"",
            trackInvocation: (tool, action, duration, success, path) =>
                invocation = new ToolInvocation(tool, action, duration, success, path));

        Assert.Equal("\"General\"", response);
        Assert.NotNull(invocation);
        Assert.True(invocation.Success);
    }

    [Fact]
    public void ExecuteToolAction_InvalidJsonResponse_TracksFailure()
    {
        ToolInvocation? invocation = null;

        var response = ExcelToolsBase.ExecuteToolAction(
            "chart",
            "get-axis-number-format",
            path: null,
            operation: () => "not-json",
            trackInvocation: (tool, action, duration, success, path) =>
                invocation = new ToolInvocation(tool, action, duration, success, path));

        Assert.Equal("not-json", response);
        Assert.NotNull(invocation);
        Assert.False(invocation.Success);
    }

    [Fact]
    public void ExecuteToolAction_ThrownException_TracksFailureWithoutTelemetryDetails()
    {
        ToolInvocation? invocation = null;

        var response = ExcelToolsBase.ExecuteToolAction(
            "range",
            "get-values",
            path: null,
            operation: () => throw new InvalidOperationException(
                @"Private failure at C:\Users\Someone\Secret.xlsx on Sheet1!A1"),
            trackInvocation: (tool, action, duration, success, path) =>
                invocation = new ToolInvocation(tool, action, duration, success, path));

        using var json = JsonDocument.Parse(response);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());
        Assert.NotNull(invocation);
        Assert.Equal(new ToolInvocation("range", "get-values", invocation.DurationMs, false, null), invocation);
    }

    private sealed record ToolInvocation(
        string Tool,
        string Action,
        long DurationMs,
        bool Success,
        string? Path);
}
