using System.Text.Json;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.McpServer.Telemetry;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "McpProtocol")]
[Trait("RequiresExcel", "false")]
public sealed class ForwardedResultConversionTests
{
    [Theory]
    [InlineData("grid")]
    [InlineData("scalar")]
    [InlineData("null")]
    [InlineData("no-result")]
    [InlineData("failure")]
    [InlineData("nested-failure")]
    [InlineData("diagnostic")]
    public async Task TypedForwarding_MatchesStringConversionAndTelemetry(string scenario)
    {
        var (tool, action, command, response) = Scenario(scenario);
        ToolInvocationResult? previousTelemetry = null;
        ToolInvocationResult? typedTelemetry = null;
        var previous = await ExcelToolsBase.ExecuteToolActionAsync(tool, action,
            () => Task.FromResult(ExcelToolsBase.ReadServiceResponse(response, command).ToToolJson()),
            CancellationToken.None, (_, _, _, invocation) => previousTelemetry = invocation);
        var typed = await ExcelToolsBase.ExecuteToolActionAsync(tool, action,
            () => Task.FromResult(ExcelToolsBase.ReadServiceResponse(response, command)),
            CancellationToken.None, (_, _, _, invocation) => typedTelemetry = invocation);

        Assert.Equal(previous.IsError, typed.IsError);
        var previousText = Assert.Single(previous.Content.OfType<TextContentBlock>()).Text;
        var typedText = Assert.Single(typed.Content.OfType<TextContentBlock>()).Text;
        using var previousJson = JsonDocument.Parse(previousText);
        using var typedJson = JsonDocument.Parse(typedText);
        Assert.True(JsonElement.DeepEquals(previousJson.RootElement, typedJson.RootElement));
        Assert.True(JsonElement.DeepEquals(typedJson.RootElement, typed.StructuredContent!.Value));
        Assert.True(JsonElement.DeepEquals(previous.StructuredContent!.Value, typed.StructuredContent.Value));
        Assert.Equal(previousTelemetry, typedTelemetry);
        Assert.Equal(scenario is "failure" or "nested-failure" or "diagnostic", typed.IsError);
        if (scenario == "failure")
        {
            Assert.Equal("Conflict", typedJson.RootElement.GetProperty("errorCategory").GetString());
            Assert.Equal("inserting", typedJson.RootElement.GetProperty("partialState").GetProperty("phase").GetString());
        }
        if (scenario == "nested-failure")
            Assert.Equal(ToolInvocationOutcome.Failed, typedTelemetry!.Value.Outcome);
        if (scenario == "diagnostic")
            Assert.Equal(ToolInvocationOutcome.ExpectedNegative, typedTelemetry!.Value.Outcome);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task TypedForwarding_CancellationBeforeOrAfterOperation_PreservesCancellation(bool afterOperation)
    {
        using var cancellation = new CancellationTokenSource();
        int operationCalls = 0;
        int telemetryCalls = 0;
        ToolInvocationResult? invocation = null;
        if (!afterOperation) cancellation.Cancel();
        Task<ResultJson> Operation()
        {
            operationCalls++;
            cancellation.Cancel();
            return Task.FromResult(ResultJson.Read("{}", "range.get-values"));
        }
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            ExcelToolsBase.ExecuteToolActionAsync("range_read", "get-values", Operation,
                cancellation.Token, (_, _, _, result) => { telemetryCalls++; invocation = result; }));
        Assert.Equal(afterOperation ? 1 : 0, operationCalls);
        Assert.Equal(1, telemetryCalls);
        Assert.Equal(ToolFailureCause.Cancellation, invocation!.Value.FailureCause);
    }

    [Theory]
    [InlineData("argument")]
    [InlineData("json")]
    [InlineData("unexpected")]
    public async Task TypedForwarding_ExceptionsKeepExistingTransportAndTelemetry(string scenario)
    {
        Exception failure = scenario switch
        {
            "argument" => new ArgumentException("Invalid range"),
            "json" => new JsonException("Malformed result"),
            _ => new InvalidOperationException("Unexpected failure")
        };
        int telemetryCalls = 0;
        ToolInvocationResult? invocation = null;
        Task<CallToolResult> Execute() => ExcelToolsBase.ExecuteToolActionAsync("range_read", "get-values",
            () => Task.FromException<ResultJson>(failure), CancellationToken.None,
            (_, _, _, result) => { telemetryCalls++; invocation = result; });
        if (scenario == "argument")
        {
            var result = await Execute();
            Assert.True(result.IsError);
            Assert.Equal("InvalidInput", result.StructuredContent!.Value.GetProperty("errorCategory").GetString());
        }
        else
        {
            var observed = await Assert.ThrowsAsync(failure.GetType(), Execute);
            Assert.Same(failure, observed);
        }
        Assert.Equal(1, telemetryCalls);
        Assert.Equal(scenario == "argument" ? ToolFailureClass.InputState :
            scenario == "json" ? ToolFailureClass.InternalProductFault : ToolFailureClass.Unclassified,
            invocation!.Value.FailureClass);
    }

    [Fact]
    public async Task TypedForwarding_ReadWithoutResultStillFailsClosed()
    {
        await Assert.ThrowsAsync<InvalidOperationException>(() =>
            ExcelToolsBase.ExecuteToolActionAsync("range_read", "get-values",
                () => Task.FromResult(ExcelToolsBase.ReadServiceResponse(
                    new ServiceResponse { Success = true }, "range.get-values")), CancellationToken.None));
    }

    private static (string Tool, string Action, string Command, ServiceResponse Response) Scenario(string scenario) => scenario switch
    {
        "no-result" => ("range", "set-values", "range.set-values", new ServiceResponse { Success = true }),
        "nested-failure" => ("powerquery", "refresh-all", "powerquery.refresh-all", new ServiceResponse
        {
            Success = true,
            Result = """{"success":true,"queryMetrics":[{"success":false,"errorCategory":"DependencyUnavailable","errorMessage":"Missing provider"}],"refreshedQueries":["First"]}"""
        }),
        "diagnostic" => ("file_read", "test", "session.test", new ServiceResponse
        {
            Success = true,
            Result = """{"canOpen":false,"exists":false,"message":"Missing fixture"}"""
        }),
        "failure" => ("range_read", "get-values", "range.get-values", new ServiceResponse
        {
            Success = false,
            Result = """{"values":[[false,0,null]],"warnings":["Check state"]}""",
            ErrorCategory = "Conflict",
            ErrorMessage = "Changed input",
            SessionId = "owned-session",
            ExceptionType = "OperationFailureException",
            HResult = "0x80004005",
            InnerError = "Original cause",
            PartialState = """{"phase":"inserting","changedCells":2}"""
        }),
        _ => ("range_read", "get-values", "range.get-values", new ServiceResponse
        {
            Success = true,
            Result = scenario switch
            {
                "scalar" => "false",
                "null" => "null",
                _ => """{"success":true,"filePath":"fixture.xlsx","values":[[false,0,null,"text"]],"cellErrors":[],"warnings":["Check cells"]}"""
            }
        })
    };
}
