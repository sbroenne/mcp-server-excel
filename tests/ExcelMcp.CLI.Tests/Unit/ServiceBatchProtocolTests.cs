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
    public void WorkflowPlanEnvelope_RoundTripsPlanCheckpointModeWithoutChangingLegacyBatch()
    {
        var request = new WorkflowPlanRequest
        {
            StopOnError = true,
            CheckpointMode = WorkflowCheckpointMode.Once,
            VerifySheetName = "Data",
            VerifyRangeAddress = "A1:A2",
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
                },
            ],
        };

        var json = ServiceProtocol.Serialize(request);
        var roundTrip = ServiceProtocol.Deserialize<WorkflowPlanRequest>(json);

        Assert.NotNull(roundTrip);
        Assert.Equal(WorkflowCheckpointMode.Once, roundTrip.CheckpointMode);
        Assert.True(roundTrip.StopOnError);
        Assert.Equal("Data", roundTrip.VerifySheetName);
        Assert.Equal("A1:A2", roundTrip.VerifyRangeAddress);
        Assert.Single(roundTrip.Operations);

        var legacy = ServiceProtocol.Deserialize<ServiceBatchRequest>(json);
        Assert.NotNull(legacy);
        Assert.True(legacy.StopOnError);
        Assert.Single(legacy.Operations);
    }

    [Fact]
    public void WorkflowPlanReceipt_RoundTripsCompactKnownAndUnknownOutcomes()
    {
        var receipt = new WorkflowPlanReceipt
        {
            PlanId = "plan-1",
            Outcome = WorkflowPlanOutcome.Completed,
            OperationCount = 2,
            AttemptedCount = 2,
            CompletedCount = 2,
            Verification = new WorkflowRangeVerificationReceipt
            {
                Status = "verified",
                SheetName = "Data",
                RangeAddress = "$A$1:$A$2",
                RowCount = 2,
                ColumnCount = 1,
                CellCount = 2,
                InspectedCellCount = 2,
                InspectedRangeAddress = "$A$1:$A$2",
                NonEmptyCellCount = 2,
                FormulaCellCount = 0,
                Fingerprint = new string('a', 64),
                Preview = [[1d], [2d]],
            },
            Steps =
            [
                new WorkflowStepReceipt { Index = 0, Command = "range.set-values", Status = "completed" },
                new WorkflowStepReceipt { Index = 1, Command = "range.set-values", Status = "completed" },
            ],
        };

        var known = ServiceProtocol.Deserialize<WorkflowPlanReceipt>(ServiceProtocol.Serialize(receipt));
        Assert.Equal(WorkflowPlanOutcome.Completed, known!.Outcome);
        Assert.Equal(2, known.Steps.Count);
        Assert.Null(known.FailedIndex);
        Assert.Equal("verified", known.Verification!.Status);
        Assert.Equal(2, known.Verification.InspectedCellCount);
        Assert.Equal(64, known.Verification.Fingerprint!.Length);
        Assert.Equal(2, known.Verification.Preview!.Count);

        var unknown = receipt with
        {
            Outcome = WorkflowPlanOutcome.Unknown,
            AttemptedCount = 1,
            CompletedCount = 1,
            FailedIndex = 1,
        };
        var unknownRoundTrip = ServiceProtocol.Deserialize<WorkflowPlanReceipt>(ServiceProtocol.Serialize(unknown));
        Assert.Equal(WorkflowPlanOutcome.Unknown, unknownRoundTrip!.Outcome);
        Assert.Equal(1, unknownRoundTrip.FailedIndex);
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
