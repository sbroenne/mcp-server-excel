using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Service.Idempotency;
using Xunit;
using ServiceBatchOperation = Sbroenne.ExcelMcp.ComInterop.ServiceClient.ServiceBatchOperation;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "Idempotency")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class IdempotencyCoordinatorTests
{
    [Fact]
    public async Task ConcurrentExactRetry_ExecutesOnceAndReturnsExactReceipt()
    {
        var coordinator = new IdempotencyCoordinator();
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var executions = 0;
        var request = CreateRequest("retry-1", "session-a", "{\"rows\":[[1]]}");

        async Task<ServiceResponse> ExecuteAsync()
        {
            Interlocked.Increment(ref executions);
            entered.TrySetResult();
            await release.Task;
            return new ServiceResponse { Success = true, Result = "{\"receipt\":\"original\"}" };
        }

        var first = coordinator.ExecuteAsync(request, ExecuteAsync);
        await entered.Task;
        var duplicate = coordinator.ExecuteAsync(request, ExecuteAsync);
        release.TrySetResult();

        var responses = await Task.WhenAll(first, duplicate);

        Assert.Equal(1, executions);
        Assert.Equal(ServiceProtocol.Serialize(responses[0]), ServiceProtocol.Serialize(responses[1]));
    }

    [Fact]
    public async Task CompletedExactRetry_ReplaysReceiptButChangedArgumentsConflict()
    {
        var coordinator = new IdempotencyCoordinator();
        var executions = 0;
        var original = CreateRequest("retry-2", "session-a", "{\"b\":2,\"a\":1}");
        var equivalent = CreateRequest("retry-2", "session-a", "{\"a\":1,\"b\":2}");
        var changed = CreateRequest("retry-2", "session-a", "{\"a\":9,\"b\":2}");

        Task<ServiceResponse> ExecuteAsync()
        {
            Interlocked.Increment(ref executions);
            return Task.FromResult(new ServiceResponse { Success = true, Result = "{\"receipt\":7}" });
        }

        var first = await coordinator.ExecuteAsync(original, ExecuteAsync);
        var replay = await coordinator.ExecuteAsync(equivalent, ExecuteAsync);
        var conflict = await coordinator.ExecuteAsync(changed, ExecuteAsync);

        Assert.Equal(1, executions);
        Assert.Equal(first.Result, replay.Result);
        Assert.False(conflict.Success);
        Assert.Equal("IdempotencyConflict", conflict.ErrorCategory);
    }

    [Fact]
    public async Task AmbiguousOutcome_IsRememberedAndNeverAutomaticallyReplayed()
    {
        var coordinator = new IdempotencyCoordinator();
        var executions = 0;
        var request = CreateRequest("retry-3", "session-a", "{\"rows\":[[3]]}");

        Task<ServiceResponse> ExecuteAsync()
        {
            Interlocked.Increment(ref executions);
            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorCategory = "Timeout",
                ErrorMessage = "The mutation may have committed."
            });
        }

        var first = await coordinator.ExecuteAsync(request, ExecuteAsync);
        var retry = await coordinator.ExecuteAsync(request, ExecuteAsync);

        Assert.Equal(1, executions);
        Assert.Equal("Timeout", first.ErrorCategory);
        Assert.Equal("IdempotencyUnknownOutcome", retry.ErrorCategory);
    }

    [Theory]
    [InlineData("TimeoutBeforeExecution")]
    [InlineData("CancelledBeforeExecution")]
    public async Task KnownNotExecutedOutcome_AllowsSameKeyToRetry(
        string firstErrorCategory)
    {
        var coordinator = new IdempotencyCoordinator();
        var executions = 0;
        var request = CreateRequest("retry-known-safe", "session-a", "{\"rows\":[[4]]}");

        Task<ServiceResponse> ExecuteAsync()
        {
            int attempt = Interlocked.Increment(ref executions);
            return Task.FromResult(attempt == 1
                ? new ServiceResponse
                {
                    Success = false,
                    ErrorCategory = firstErrorCategory,
                    ErrorMessage = "The delegate never started."
                }
                : new ServiceResponse { Success = true, Result = "{\"receipt\":\"retry\"}" });
        }

        var first = await coordinator.ExecuteAsync(request, ExecuteAsync);
        var retry = await coordinator.ExecuteAsync(request, ExecuteAsync);

        Assert.Equal(firstErrorCategory, first.ErrorCategory);
        Assert.True(retry.Success);
        Assert.Equal(2, executions);
    }

    [Fact]
    public async Task CheckpointFailedOutcome_AllowsSameKeyToRetry()
    {
        var coordinator = new IdempotencyCoordinator();
        var executions = 0;
        var request = CreateRequest("retry-checkpoint", "session-a", "{\"rows\":[[5]]}");

        Task<ServiceResponse> ExecuteAsync()
        {
            int attempt = Interlocked.Increment(ref executions);
            return Task.FromResult(attempt == 1
                ? new ServiceResponse { Success = false, ErrorCategory = "CheckpointFailed", ErrorMessage = "checkpoint unavailable" }
                : new ServiceResponse { Success = true, Result = "{\"receipt\":\"after-checkpoint\"}" });
        }

        var first = await coordinator.ExecuteAsync(request, ExecuteAsync);
        var retry = await coordinator.ExecuteAsync(request, ExecuteAsync);

        Assert.Equal("CheckpointFailed", first.ErrorCategory);
        Assert.True(retry.Success);
        Assert.Equal(2, executions);
    }

    [Fact]
    public async Task ExactRetryFromDifferentSource_ReplaysReceipt()
    {
        var coordinator = new IdempotencyCoordinator();
        var executions = 0;
        var cliRequest = CreateRequest("cross-source", "session-a", "{\"rows\":[[6]]}");
        var mcpRequest = new ServiceRequest
        {
            Command = cliRequest.Command,
            SessionId = cliRequest.SessionId,
            Args = cliRequest.Args,
            ReviewId = cliRequest.ReviewId,
            Checkpoint = cliRequest.Checkpoint,
            Source = "mcp",
            IdempotencyKey = cliRequest.IdempotencyKey,
        };

        Task<ServiceResponse> ExecuteAsync()
        {
            Interlocked.Increment(ref executions);
            return Task.FromResult(new ServiceResponse { Success = true, Result = "{\"receipt\":\"shared\"}" });
        }

        var first = await coordinator.ExecuteAsync(cliRequest, ExecuteAsync);
        var replay = await coordinator.ExecuteAsync(mcpRequest, ExecuteAsync);

        Assert.True(first.Success);
        Assert.True(replay.Success);
        Assert.Equal(first.Result, replay.Result);
        Assert.Equal(1, executions);
    }

    [Fact]
    public async Task PendingExactRetry_ReturnsInProgressAfterBoundedWait()
    {
        var coordinator = new IdempotencyCoordinator(pendingWaitTimeout: TimeSpan.FromMilliseconds(20));
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var request = CreateRequest("pending-timeout", "session-a", "{\"rows\":[[7]]}");

        async Task<ServiceResponse> ExecuteAsync()
        {
            entered.TrySetResult();
            await release.Task;
            return new ServiceResponse { Success = true };
        }

        var first = coordinator.ExecuteAsync(request, ExecuteAsync);
        await entered.Task;
        var retry = await coordinator.ExecuteAsync(request, ExecuteAsync);

        Assert.False(retry.Success);
        Assert.Equal("IdempotencyInProgress", retry.ErrorCategory);
        release.TrySetResult();
        await first;
    }

    [Fact]
    public async Task CapacityExceeded_DoesNotEvictPendingEntry()
    {
        var coordinator = new IdempotencyCoordinator(maximumEntries: 1);
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var firstRequest = CreateRequest("capacity-1", "session-a", "{\"rows\":[[8]]}");
        var secondRequest = CreateRequest("capacity-2", "session-a", "{\"rows\":[[9]]}");

        async Task<ServiceResponse> ExecuteAsync()
        {
            entered.TrySetResult();
            await release.Task;
            return new ServiceResponse { Success = true };
        }

        var first = coordinator.ExecuteAsync(firstRequest, ExecuteAsync);
        await entered.Task;
        var second = await coordinator.ExecuteAsync(secondRequest, ExecuteAsync);

        Assert.False(second.Success);
        Assert.Equal("IdempotencyCapacityExceeded", second.ErrorCategory);
        release.TrySetResult();
        await first;
    }

    [Fact]
    public async Task SameKeyInDifferentSession_FailsClosedWithoutExecution()
    {
        var coordinator = new IdempotencyCoordinator();
        var executions = 0;

        Task<ServiceResponse> ExecuteAsync()
        {
            Interlocked.Increment(ref executions);
            return Task.FromResult(new ServiceResponse { Success = true, Result = "{}" });
        }

        _ = await coordinator.ExecuteAsync(CreateRequest("shared-key", "session-a", "{}"), ExecuteAsync);
        var conflict = await coordinator.ExecuteAsync(CreateRequest("shared-key", "session-b", "{}"), ExecuteAsync);

        Assert.Equal(1, executions);
        Assert.False(conflict.Success);
        Assert.Equal("IdempotencyScopeConflict", conflict.ErrorCategory);
    }

    [Fact]
    public async Task WorkflowPlan_ExactRetryReplaysReceiptAndChangedPlanConflicts()
    {
        var coordinator = new IdempotencyCoordinator();
        var executions = 0;
        var original = CreateWorkflowPlanRequest("workflow-retry", "{\"operations\":[{\"command\":\"range.set-values\",\"args\":{\"rangeAddress\":\"A1\",\"values\":[[1]]}}],\"checkpointMode\":\"once\"}");
        var equivalent = CreateWorkflowPlanRequest("workflow-retry", "{\"checkpointMode\":\"once\",\"operations\":[{\"command\":\"range.set-values\",\"args\":{\"values\":[[1]],\"rangeAddress\":\"A1\"}}]}");
        var changed = CreateWorkflowPlanRequest("workflow-retry", "{\"operations\":[{\"command\":\"range.set-values\",\"args\":{\"rangeAddress\":\"A1\",\"values\":[[2]]}}],\"checkpointMode\":\"once\"}");

        Task<ServiceResponse> ExecuteAsync()
        {
            Interlocked.Increment(ref executions);
            return Task.FromResult(new ServiceResponse { Success = true, Result = "{\"outcome\":\"completed\"}" });
        }

        var first = await coordinator.ExecuteAsync(original, ExecuteAsync);
        var replay = await coordinator.ExecuteAsync(equivalent, ExecuteAsync);
        var conflict = await coordinator.ExecuteAsync(changed, ExecuteAsync);

        Assert.Equal(1, executions);
        Assert.Equal(first.Result, replay.Result);
        Assert.Equal("IdempotencyConflict", conflict.ErrorCategory);
    }

    [Fact]
    public async Task WorkflowPlan_UnknownOutcomeReturnsReceiptButNeverDispatchesRetry()
    {
        var coordinator = new IdempotencyCoordinator();
        var executions = 0;
        var request = CreateWorkflowPlanRequest("workflow-unknown", "{\"operations\":[{\"command\":\"range.set-values\"}],\"checkpointMode\":\"once\"}");

        Task<ServiceResponse> ExecuteAsync()
        {
            Interlocked.Increment(ref executions);
            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorCategory = "UnknownOutcome",
                Result = "{\"outcome\":\"unknown\",\"failedIndex\":0}"
            });
        }

        var first = await coordinator.ExecuteAsync(request, ExecuteAsync);
        var retry = await coordinator.ExecuteAsync(request, ExecuteAsync);

        Assert.Equal(1, executions);
        Assert.Equal("UnknownOutcome", first.ErrorCategory);
        Assert.Equal("IdempotencyUnknownOutcome", retry.ErrorCategory);
        Assert.Equal(first.Result, retry.Result);
    }

    [Fact]
    public async Task WorkflowPlan_ReviewUnavailableBeforeExecution_AllowsSameKeyToRetry()
    {
        var coordinator = new IdempotencyCoordinator();
        var executions = 0;
        var request = CreateWorkflowPlanRequest(
            "workflow-review-unavailable",
            "{\"operations\":[{\"command\":\"range.set-values\"}]}");

        Task<ServiceResponse> ExecuteAsync()
        {
            int attempt = Interlocked.Increment(ref executions);
            return Task.FromResult(attempt == 1
                ? new ServiceResponse
                {
                    Success = false,
                    ErrorCategory = "PlanReviewUnavailable",
                    ErrorMessage = "Plan review is unavailable."
                }
                : new ServiceResponse { Success = true, Result = "{\"outcome\":\"completed\"}" });
        }

        var first = await coordinator.ExecuteAsync(request, ExecuteAsync);
        var retry = await coordinator.ExecuteAsync(request, ExecuteAsync);

        Assert.Equal("PlanReviewUnavailable", first.ErrorCategory);
        Assert.True(retry.Success);
        Assert.Equal(2, executions);
    }

    [Fact]
    public async Task WorkflowPlan_PartialKnownFailure_ReplaysReceiptWithoutDispatchingAgain()
    {
        var coordinator = new IdempotencyCoordinator();
        var executions = 0;
        var request = CreateWorkflowPlanRequest(
            "workflow-partial-failure",
            "{\"operations\":[{\"command\":\"range.set-values\"},{\"command\":\"range.set-values\"}]}");

        Task<ServiceResponse> ExecuteAsync()
        {
            Interlocked.Increment(ref executions);
            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorCategory = "PlanFailed",
                ErrorMessage = "A later step failed after an earlier mutation.",
                Result = "{\"outcome\":\"failed\",\"completedCount\":1,\"failedIndex\":1}"
            });
        }

        var first = await coordinator.ExecuteAsync(request, ExecuteAsync);
        var retry = await coordinator.ExecuteAsync(request, ExecuteAsync);

        Assert.Equal(1, executions);
        Assert.Equal("PlanFailed", first.ErrorCategory);
        Assert.Equal(ServiceProtocol.Serialize(first), ServiceProtocol.Serialize(retry));
    }

    [Fact]
    public void ProtocolRoundTrip_PreservesKeysAndOldRequestsRemainCompatible()
    {
        var request = CreateRequest("request-key", "session-a", "{}");
        var operation = new ServiceBatchOperation
        {
            Command = "table.append",
            SessionId = "session-a",
            IdempotencyKey = "operation-key"
        };

        var requestRoundTrip = ServiceProtocol.Deserialize<ServiceRequest>(ServiceProtocol.Serialize(request));
        var operationRoundTrip = ServiceProtocol.Deserialize<ServiceBatchOperation>(ServiceProtocol.Serialize(operation));
        var oldRequest = ServiceProtocol.Deserialize<ServiceRequest>("{\"command\":\"table.append\"}");

        Assert.Equal("request-key", requestRoundTrip!.IdempotencyKey);
        Assert.Equal("operation-key", operationRoundTrip!.IdempotencyKey);
        Assert.Null(oldRequest!.IdempotencyKey);
    }

    private static ServiceRequest CreateRequest(string key, string sessionId, string args) => new()
    {
        Command = "table.append",
        SessionId = sessionId,
        Args = args,
        ReviewId = "review-1",
        Source = "test",
        IdempotencyKey = key
    };

    private static ServiceRequest CreateWorkflowPlanRequest(string key, string args) => new()
    {
        Command = "workflow.execute-plan",
        SessionId = "session-workflow",
        Args = args,
        Source = "test",
        IdempotencyKey = key
    };
}
