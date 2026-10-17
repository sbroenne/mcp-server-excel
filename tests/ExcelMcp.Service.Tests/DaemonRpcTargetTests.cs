using Nerdbank.Streams;
using Sbroenne.ExcelMcp.Service.Rpc;
using StreamJsonRpc;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "StreamJsonRpc")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class DaemonRpcTargetTests
{
    [Fact]
    public async Task ProcessCommandAsync_CancelledStartup_ObservesOriginalLateResponse()
    {
        var original = new TaskCompletionSource<ServiceResponse>(TaskCreationOptions.RunContinuationsAsynchronously);
        var abandoned = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var late = new TaskCompletionSource<ServiceResponse>(TaskCreationOptions.RunContinuationsAsynchronously);
        var request = new ServiceRequest { Command = "session.open", Args = """{"filePath":"retained.xlsx"}""" };
        var target = new DaemonRpcTarget(_ => original.Task, static () => { }, (received, response) =>
        {
            Assert.Same(request, received);
            if (response is null)
                abandoned.TrySetResult();
            else
                late.TrySetResult(response);
        });
        using var cancellation = new CancellationTokenSource();
        var execution = target.ProcessCommandAsync(request, cancellation.Token);
        cancellation.Cancel();
        await abandoned.Task.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.False(execution.IsCompleted);
        Assert.False(original.Task.IsCompleted);

        var expected = new ServiceResponse { Success = true, Result = """{"sessionId":"retained-session"}""" };
        original.SetResult(expected);
        Assert.Same(expected, await execution.WaitAsync(TimeSpan.FromSeconds(5)));
        Assert.Same(expected, await late.Task.WaitAsync(TimeSpan.FromSeconds(5)));
    }

    [Fact]
    public async Task ProcessCommandAsync_DelegatesExactRequestAndRecordsActivity()
    {
        ServiceRequest? capturedRequest = null;
        var activityCount = 0;
        var expectedResponse = new ServiceResponse
        {
            Success = true,
            Result = """{"value":42}"""
        };
        var target = new DaemonRpcTarget(
            request =>
            {
                capturedRequest = request;
                return Task.FromResult(expectedResponse);
            },
            () => activityCount++);
        var request = new ServiceRequest
        {
            Command = "diag.echo",
            SessionId = "session-1",
            Args = """{"message":"hello"}"""
        };

        var response = await RoundTripAsync(target, request);

        Assert.NotNull(capturedRequest);
        Assert.Equal(request.Command, capturedRequest.Command);
        Assert.Equal(request.SessionId, capturedRequest.SessionId);
        Assert.Equal(request.Args, capturedRequest.Args);
        Assert.Equal(1, activityCount);
        Assert.True(response.Success);
        Assert.Equal(expectedResponse.Result, response.Result);
    }

    [Fact]
    public async Task ProcessCommandAsync_ErrorResponseRoundTripsAsData()
    {
        var target = new DaemonRpcTarget(
            _ => Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = "Controlled failure"
            }),
            static () => { });

        var response = await RoundTripAsync(
            target,
            new ServiceRequest { Command = "controlled.failure" });

        Assert.False(response.Success);
        Assert.Equal("InvalidInput", response.ErrorCategory);
        Assert.Equal("Controlled failure", response.ErrorMessage);
    }

    [Fact]
    public async Task ServerRpc_Completion_ResolvesWhenClientDisconnects()
    {
        var (serverStream, clientStream) = FullDuplexStream.CreatePair();
        var target = new DaemonRpcTarget(
            _ => Task.FromResult(new ServiceResponse { Success = true }),
            static () => { });
        using var serverRpc = JsonRpc.Attach(serverStream, target);
        var clientProxy = JsonRpc.Attach<IExcelDaemonRpc>(clientStream);

        ((IDisposable)clientProxy).Dispose();

        var completedTask = await Task.WhenAny(
            serverRpc.Completion,
            Task.Delay(TimeSpan.FromSeconds(5)));
        Assert.Same(serverRpc.Completion, completedTask);
    }

    private static async Task<ServiceResponse> RoundTripAsync(
        DaemonRpcTarget target,
        ServiceRequest request)
    {
        var (serverStream, clientStream) = FullDuplexStream.CreatePair();
        using var serverRpc = JsonRpc.Attach(serverStream, target);
        var clientProxy = JsonRpc.Attach<IExcelDaemonRpc>(clientStream);
        try
        {
            return await clientProxy.ProcessCommandAsync(request);
        }
        finally
        {
            ((IDisposable)clientProxy).Dispose();
        }
    }
}
