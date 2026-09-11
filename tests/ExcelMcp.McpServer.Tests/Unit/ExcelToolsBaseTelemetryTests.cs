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
    [Theory]
    [InlineData("wrapped-com", "ComInterop", "ExcelRuntime")]
    [InlineData("json", "InvalidInput", "InputState")]
    [InlineData("query", "Syntax", "InputState")]
    [InlineData("prerequisite", "Prerequisite", "InputState")]
    [InlineData("dependency", "DependencyUnavailable", "ExternalDependency")]
    [InlineData("permissions", "Permissions", "ExternalDependency")]
    public void ExecuteToolAction_KnownException_PreservesCategory(
        string scenario, string category, string failureClass)
    {
#pragma warning disable CA2201 // Synthetic exceptions exercise serialization, not COM behavior.
        Exception error = scenario switch
        {
            "wrapped-com" => new InvalidOperationException("Operation context",
                new System.Reflection.TargetInvocationException(
                    new System.Runtime.InteropServices.COMException("Excel failure", unchecked((int)0x800A03EC)))),
            "json" => new JsonException("Invalid arguments"),
            "prerequisite" => new OperationFailureException(OperationFailureCategory.Prerequisite, "Missing model"),
            "dependency" => new OperationFailureException(OperationFailureCategory.DependencyUnavailable, "Missing provider"),
            "permissions" => new OperationFailureException(OperationFailureCategory.Permissions, "Access blocked"),
            _ => new Sbroenne.ExcelMcp.Core.Commands.PowerQueryCommandException(
                "Invalid query", "Syntax", new InvalidOperationException("Query details"))
        };
#pragma warning restore CA2201
        ToolInvocationResult? invocation = null;
        var response = ExcelToolsBase.ExecuteToolAction("powerquery", "evaluate", null,
            () => throw error, (_, _, _, result) => invocation = result);

        using var document = JsonDocument.Parse(response);
        Assert.Equal(category, document.RootElement.GetProperty("errorCategory").GetString());
        Assert.Contains(error.Message, document.RootElement.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
        Assert.Equal(new ToolInvocationResult(ToolInvocationOutcome.Failed,
            Enum.Parse<ToolFailureClass>(failureClass)), invocation);
        if (scenario == "wrapped-com")
        {
            Assert.Equal("0x800A03EC", document.RootElement.GetProperty("hresult").GetString());
        }
    }

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
    [InlineData("Prerequisite", "InputState")]
    [InlineData("DependencyUnavailable", "ExternalDependency")]
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

    [Theory]
    [InlineData("Unclassified", "unclassified")]
    [InlineData("InputState", "input-state")]
    [InlineData("ExternalDependency", "external-dependency")]
    public void CreateToolInvocationTelemetry_FailureEmitsOnlyAllowlistedClassification(
        string failureClass, string label)
    {
        var result = new ToolInvocationResult(
            ToolInvocationOutcome.Failed,
            Enum.Parse<ToolFailureClass>(failureClass));

        var (eventTelemetry, requestTelemetry) =
            ExcelMcpTelemetry.CreateToolInvocationTelemetry(
                "range",
                "get-values",
                25,
                result);

        Assert.False(requestTelemetry.Success);
        Assert.Equal("500", requestTelemetry.ResponseCode);
        Assert.Equal("failed", requestTelemetry.Properties["Outcome"]);
        Assert.Equal(label, requestTelemetry.Properties["FailureClass"]);
        Assert.Equal("failed", eventTelemetry.Properties["Outcome"]);
        Assert.Equal(label, eventTelemetry.Properties["FailureClass"]);

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
