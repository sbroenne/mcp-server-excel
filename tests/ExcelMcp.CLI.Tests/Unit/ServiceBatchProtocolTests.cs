using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Commands;
using Sbroenne.ExcelMcp.ComInterop.ServiceClient;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "Batch")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ServiceBatchProtocolTests
{
    [Fact]
    public void BatchEnvelope_RoundTripsOperationsAndSafetyContext()
    {
        var request = new ServiceBatchRequest
        {
            StopOnError = false,
            Operations =
            [
                new ServiceBatchOperation
                {
                    Command = "range.set-values",
                    Args = JsonSerializer.SerializeToElement(new
                    {
                        sheetName = "Data",
                        rangeAddress = "A1",
                        values = new object?[][] { [1] },
                    }, ServiceProtocol.JsonOptions),
                    ReviewOnly = true,
                    ReviewId = "review-1",
                    Checkpoint = true,
                },
            ],
        };

        var json = ServiceProtocol.Serialize(request);
        var roundTrip = ServiceProtocol.Deserialize<ServiceBatchRequest>(json);

        Assert.NotNull(roundTrip);
        Assert.False(roundTrip.StopOnError);
        var operation = Assert.Single(roundTrip.Operations);
        Assert.Equal("range.set-values", operation.Command);
        Assert.True(operation.ReviewOnly);
        Assert.Equal("review-1", operation.ReviewId);
        Assert.True(operation.Checkpoint);
    }

    [Fact]
    public void BatchEnvelope_WhenStopOnErrorIsOmitted_DefaultsToFailClosed()
    {
        const string json = "{\"operations\":[{\"command\":\"range.get-values\"}]}";

        var request = ServiceProtocol.Deserialize<ServiceBatchRequest>(json);

        Assert.NotNull(request);
        Assert.True(request.StopOnError);
    }

    [Fact]
    public void ServerBatchLimit_RejectsCommandsBeyondServiceCapacity()
    {
        Assert.True(BatchCommand.IsWithinServerBatchLimit(256));
        Assert.False(BatchCommand.IsWithinServerBatchLimit(257));
    }

    [Fact]
    public void ServerBatchResponse_RejectsMalformedEnvelopes()
    {
#pragma warning disable CS8618
        var nullResults = new ServiceBatchResponse { Results = null! };
#pragma warning restore CS8618
        Assert.False(BatchCommand.TryValidateServerBatchResponse(null, 2, out _));
        Assert.False(BatchCommand.TryValidateServerBatchResponse(nullResults, 2, out _));
        Assert.False(BatchCommand.TryValidateServerBatchResponse(
            new ServiceBatchResponse
            {
                Completed = true,
                Results =
                [
                    new ServiceBatchOperationResult { Index = 1, Success = true },
                ],
            },
            2,
            out _));
        Assert.False(BatchCommand.TryValidateServerBatchResponse(
            new ServiceBatchResponse
            {
                Success = false,
                Completed = true,
                Results =
                [
                    new ServiceBatchOperationResult { Index = 0, Success = true },
                ],
            },
            1,
            out _));
    }

    [Fact]
    public void ServerBatchResponse_AllowsPreExecutionValidationFailureAtOriginalIndex()
    {
        var response = new ServiceBatchResponse
        {
            Success = false,
            Completed = false,
            FailedIndex = 2,
            Results =
            [
                new ServiceBatchOperationResult
                {
                    Index = 2,
                    Success = false,
                    ErrorCategory = "InvalidInput",
                    ErrorMessage = "invalid command",
                },
            ],
        };

        Assert.True(BatchCommand.TryValidateServerBatchResponse(response, 3, out _));
    }

    [Fact]
    public void ServerBatchResponse_PreservesStructuredOperationErrors()
    {
        var response = new ServiceBatchResponse
        {
            Results =
            [
                new ServiceBatchOperationResult
                {
                    Index = 0,
                    Success = false,
                    ErrorMessage = "not executed",
                    ErrorCategory = "TimeoutBeforeExecution",
                    ExceptionType = "ExcelOperationNotStartedTimeoutException",
                    HResult = "0x80004005",
                },
            ],
        };

        var roundTrip = ServiceProtocol.Deserialize<ServiceBatchResponse>(ServiceProtocol.Serialize(response));
        var result = Assert.Single(roundTrip!.Results);
        Assert.Equal("TimeoutBeforeExecution", result.ErrorCategory);
        Assert.Equal("ExcelOperationNotStartedTimeoutException", result.ExceptionType);
        Assert.Equal("0x80004005", result.HResult);
    }
}
