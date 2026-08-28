using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.McpServer.Telemetry;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Sbroenne.ExcelMcp.Generated;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Telemetry")]
public sealed class ExcelToolsBaseTelemetryTests
{
    [Fact]
    public void ExecuteToolAction_SuccessResponse_TracksSucceeded()
    {
        ToolInvocationResult? invocation = null;

        var response = Execute(
            """{"success":true,"value":"private workbook content"}""",
            result => invocation = result);

        Assert.Equal("""{"success":true,"value":"private workbook content"}""", response);
        Assert.Equal(
            new ToolInvocationResult(ToolInvocationOutcome.Succeeded, null),
            invocation);
    }

    [Fact]
    public void ExecuteToolAction_DiagnosticNegative_TracksExpectedNegative()
    {
        ToolInvocationResult? invocation = null;
        var diagnostic = JsonSerializer.Serialize(
            new FileValidationInfo
            {
                CanOpen = false,
                Exists = false,
                Message = "private diagnostic detail"
            },
            ExcelToolsBase.JsonOptions);

        var response = Execute(
            diagnostic,
            result => invocation = result,
            toolName: "file",
            actionName: "test");

        Assert.Equal(diagnostic, response);
        Assert.DoesNotContain("\"isError\"", response, StringComparison.Ordinal);
        Assert.Equal(
            new ToolInvocationResult(ToolInvocationOutcome.ExpectedNegative, null),
            invocation);
    }

    [Theory]
    [InlineData("InvalidInput", "InputState")]
    [InlineData("SessionNotFound", "InputState")]
    [InlineData("Privacy", "ExternalDependency")]
    [InlineData("Timeout", "TimeoutCancellation")]
    [InlineData("ComInterop", "ExcelRuntime")]
    [InlineData("ServiceStartup", "InternalProductFault")]
    [InlineData("FutureCategory", "Unclassified")]
    public void ExecuteToolAction_StructuredFailure_UsesAllowlistedClass(
        string errorCategory,
        string expectedFailureClass)
    {
        ToolInvocationResult? invocation = null;
        var json = JsonSerializer.Serialize(new
        {
            success = false,
            isError = true,
            errorCategory,
            errorMessage = @"Private detail at C:\Users\Someone\Secret.xlsx"
        });

        Execute(json, result => invocation = result);

        Assert.Equal(
            new ToolInvocationResult(
                ToolInvocationOutcome.Failed,
                Enum.Parse<ToolFailureClass>(expectedFailureClass)),
            invocation);
    }

    [Fact]
    public void ExecuteToolAction_FailureWithoutCategory_TracksUnclassified()
    {
        ToolInvocationResult? invocation = null;

        Execute(
            """{"success":false,"isError":true,"errorMessage":"private detail"}""",
            result => invocation = result);

        Assert.Equal(
            new ToolInvocationResult(
                ToolInvocationOutcome.Failed,
                ToolFailureClass.Unclassified),
            invocation);
    }

    [Fact]
    public void ExecuteToolAction_NegativeCoreResultWithoutIsError_TracksFailure()
    {
        ToolInvocationResult? invocation = null;

        Execute(
            """{"success":false,"errorMessage":"private command failure"}""",
            result => invocation = result);

        Assert.Equal(
            new ToolInvocationResult(
                ToolInvocationOutcome.Failed,
                ToolFailureClass.Unclassified),
            invocation);
    }

    [Fact]
    public void ExecuteToolAction_PrimitiveJsonResponse_TracksSucceeded()
    {
        ToolInvocationResult? invocation = null;

        var response = Execute("\"General\"", result => invocation = result);

        Assert.Equal("\"General\"", response);
        Assert.Equal(
            new ToolInvocationResult(ToolInvocationOutcome.Succeeded, null),
            invocation);
    }

    [Fact]
    public void ExecuteToolAction_InvalidJsonResponse_TracksUnclassifiedFailure()
    {
        ToolInvocationResult? invocation = null;

        var response = Execute("not-json", result => invocation = result);

        Assert.Equal("not-json", response);
        Assert.Equal(
            new ToolInvocationResult(
                ToolInvocationOutcome.Failed,
                ToolFailureClass.Unclassified),
            invocation);
    }

    [Fact]
    public void ExecuteToolAction_ThrownException_TracksUnclassifiedFailure()
    {
        ToolInvocationResult? invocation = null;

        var response = ExcelToolsBase.ExecuteToolAction(
            "range",
            "get-values",
            path: null,
            operation: () => throw new InvalidOperationException(
                @"Private failure at C:\Users\Someone\Secret.xlsx on Sheet1!A1"),
            trackInvocation: (_, _, _, result) => invocation = result);

        using var json = JsonDocument.Parse(response);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(
            new ToolInvocationResult(
                ToolInvocationOutcome.Failed,
                ToolFailureClass.Unclassified),
            invocation);
    }

    [Fact]
    public void CreateToolInvocationTelemetry_ExpectedNegativeIsSuccessfulRequest()
    {
        var result = new ToolInvocationResult(ToolInvocationOutcome.ExpectedNegative, null);

        var (eventTelemetry, requestTelemetry) =
            ExcelMcpTelemetry.CreateToolInvocationTelemetry(
                "file",
                "test",
                12,
                result);

        Assert.True(requestTelemetry.Success);
        Assert.Equal("200", requestTelemetry.ResponseCode);
        Assert.Equal("expected-negative", requestTelemetry.Properties["Outcome"]);
        Assert.False(requestTelemetry.Properties.ContainsKey("FailureClass"));
        Assert.Equal("expected-negative", eventTelemetry.Properties["Outcome"]);
    }

    [Fact]
    public void CreateToolInvocationTelemetry_FailureEmitsOnlyAllowlistedClassification()
    {
        var result = new ToolInvocationResult(
            ToolInvocationOutcome.Failed,
            ToolFailureClass.Unclassified);

        var (eventTelemetry, requestTelemetry) =
            ExcelMcpTelemetry.CreateToolInvocationTelemetry(
                "range",
                "get-values",
                25,
                result);

        Assert.False(requestTelemetry.Success);
        Assert.Equal("500", requestTelemetry.ResponseCode);
        Assert.Equal("failed", requestTelemetry.Properties["Outcome"]);
        Assert.Equal("unclassified", requestTelemetry.Properties["FailureClass"]);
        Assert.Equal("failed", eventTelemetry.Properties["Outcome"]);
        Assert.Equal("unclassified", eventTelemetry.Properties["FailureClass"]);

        var serializedProperties = string.Join(
            "\n",
            eventTelemetry.Properties.Concat(requestTelemetry.Properties));
        Assert.DoesNotContain("Secret.xlsx", serializedProperties, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("errorMessage", serializedProperties, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Exception", serializedProperties, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("FileSessionId", serializedProperties, StringComparison.Ordinal);
    }

    [Fact]
    public void WorksheetMissingSession_ReturnsCategorizedRecoveryGuidance()
    {
        var response = ExcelWorksheetTool.ExcelWorksheet(SheetAction.List);

        using var json = JsonDocument.Parse(response);
        var root = json.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean());
        Assert.True(root.GetProperty("isError").GetBoolean());
        Assert.Equal("InvalidInput", root.GetProperty("errorCategory").GetString());
        Assert.Contains("file 'open'", root.GetProperty("errorMessage").GetString());
    }

    private static string Execute(
        string response,
        Action<ToolInvocationResult> capture,
        string toolName = "range",
        string actionName = "get-values") =>
        ExcelToolsBase.ExecuteToolAction(
            toolName,
            actionName,
            path: null,
            operation: () => response,
            trackInvocation: (_, _, _, result) => capture(result));
}
