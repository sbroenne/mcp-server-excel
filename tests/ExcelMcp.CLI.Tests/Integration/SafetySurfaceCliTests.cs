using System.IO.Pipes;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using StreamJsonRpc;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Verifies generated CLI mutation commands expose the shared safety handshake.
/// </summary>
[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "Safety")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class SafetySurfaceCliTests
{
    [Fact]
    public async Task RangeHelp_ExposesReviewAndCheckpointOptions()
    {
        var result = await CliProcessHelper.RunAsync("range --help");

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("--review-only", result.Stdout, StringComparison.Ordinal);
        Assert.Contains("--review-id", result.Stdout, StringComparison.Ordinal);
        Assert.Contains("--checkpoint", result.Stdout, StringComparison.Ordinal);
        Assert.Contains("--idempotency-key", result.Stdout, StringComparison.Ordinal);
    }

    [Fact]
    public async Task GeneratedRangeCommand_ForwardsReviewAndCheckpointOptions()
    {
        var pipeName = $"excelmcp-safety-cli-{Guid.NewGuid():N}";
        await using var daemon = new CapturingDaemon(pipeName);
        await daemon.StartAsync();

        var result = await CliProcessHelper.RunAsync(
            [
                "range", "set-values",
                "--session", "safety-session",
                "--sheet", "Sheet1",
                "--range", "A1",
                "--values", "[[42]]",
                "--review-only",
                "--review-id", "review-123",
                "--checkpoint",
                "--idempotency-key", "cli-retry-123"
            ],
            environmentVariables: new Dictionary<string, string> { ["EXCELMCP_CLI_PIPE"] = pipeName });

        var request = await daemon.CapturedRequest.WaitAsync(TimeSpan.FromSeconds(10));

        Assert.Equal(0, result.ExitCode);
        Assert.Equal("range.set-values", request.Command);
        Assert.Equal("cli", request.Source);
        Assert.True(request.ReviewOnly);
        Assert.Equal("review-123", request.ReviewId);
        Assert.True(request.Checkpoint);
        Assert.Equal("cli-retry-123", request.IdempotencyKey);
    }

    [Fact]
    public async Task GeneratedSheetAtomicAction_WithoutSession_RoutesToService()
    {
        var request = await RunAndCaptureAsync(
            ["sheet", "copy-to-file", "--source-file", "source.xlsx", "--source-sheet", "Source", "--target-file", "target.xlsx"]);

        Assert.Equal("sheet.copy-to-file", request.Command);
        Assert.True(string.IsNullOrWhiteSpace(request.SessionId));
        using var args = System.Text.Json.JsonDocument.Parse(request.Args!);
        Assert.Equal("source.xlsx", args.RootElement.GetProperty("sourceFile").GetString());
        Assert.Equal("Source", args.RootElement.GetProperty("sourceSheet").GetString());
        Assert.Equal("target.xlsx", args.RootElement.GetProperty("targetFile").GetString());
    }

    [Theory]
    [InlineData("--review-only")]
    [InlineData("--checkpoint")]
    public async Task GeneratedSheetAtomicSafetyOption_WithoutSession_ReachesServiceAndFailsClosed(string safetyOption)
    {
        var pipeName = $"excelmcp-safety-cli-{Guid.NewGuid():N}";
        var stateRoot = Path.Combine(Path.GetTempPath(), $"excelmcp-atomic-cli-safety-{Guid.NewGuid():N}");
        using var service = new ExcelMcpService(stateRoot);
        await using var daemon = new CapturingDaemon(pipeName, service.ProcessAsync);
        await daemon.StartAsync();

        try
        {
            var (result, json) = await CliProcessHelper.RunJsonAsync(
                ["sheet", "copy-to-file", "--source-file", "source.xlsx", "--source-sheet", "Source", "--target-file", "target.xlsx", safetyOption],
                environmentVariables: new Dictionary<string, string> { ["EXCELMCP_CLI_PIPE"] = pipeName });
            using (json)
            {
                Assert.Equal(1, result.ExitCode);
                Assert.Equal("SafetyWorkflowUnavailable", json.RootElement.GetProperty("errorCategory").GetString());
            }

            var request = await daemon.CapturedRequest.WaitAsync(TimeSpan.FromSeconds(10));
            Assert.True(string.IsNullOrWhiteSpace(request.SessionId));
            Assert.Equal("sheet.copy-to-file", request.Command);
        }
        finally
        {
            if (Directory.Exists(stateRoot))
            {
                Directory.Delete(stateRoot, recursive: true);
            }
        }
    }

    [Fact]
    public async Task GeneratedSheetOrdinaryAction_WithoutSession_FailsBeforeServiceRouting()
    {
        var result = await CliProcessHelper.RunAsync(["sheet", "list"]);

        Assert.Equal(1, result.ExitCode);
        using var json = System.Text.Json.JsonDocument.Parse(result.Stdout);
        Assert.Contains("Session ID is required", json.RootElement.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task SafetySessionCommands_ForwardExactServiceContracts()
    {
        var configure = await RunAndCaptureAsync(
            ["session", "configure-safety", "--session-id", "safety-session", "--review-mode", "required", "--checkpoint-mode", "onRequest", "--journal-mode", "on", "--verification-mode", "on", "--abnormal-shutdown-policy", "discardWithRecoveryEvidence"]);
        Assert.Equal("session.configure-safety", configure.Command);
        Assert.Equal("safety-session", configure.SessionId);
        using (var configureArgs = System.Text.Json.JsonDocument.Parse(configure.Args!))
        {
            var root = configureArgs.RootElement;
            Assert.Equal("required", root.GetProperty("reviewMode").GetString());
            Assert.Equal("onRequest", root.GetProperty("checkpointMode").GetString());
            Assert.Equal("on", root.GetProperty("journalMode").GetString());
            Assert.Equal("on", root.GetProperty("verificationMode").GetString());
            Assert.Equal("discardWithRecoveryEvidence", root.GetProperty("abnormalShutdownPolicy").GetString());
        }

        var journal = await RunAndCaptureAsync(["session", "journal", "--session-id", "safety-session"]);
        Assert.Equal("session.journal", journal.Command);
        Assert.Equal("safety-session", journal.SessionId);
        Assert.Null(journal.Args);

        var recoveries = await RunAndCaptureAsync(["session", "recoveries"]);
        Assert.Equal("recovery.list", recoveries.Command);
        Assert.Null(recoveries.SessionId);
        Assert.Null(recoveries.Args);

        var recover = await RunAndCaptureAsync(
            ["session", "recover", "--recovery-id", "recovery-123", "--show", "--timeout", "90"]);
        Assert.Equal("recovery.recover", recover.Command);
        Assert.Null(recover.SessionId);
        using var recoverArgs = System.Text.Json.JsonDocument.Parse(recover.Args!);
        Assert.Equal("recovery-123", recoverArgs.RootElement.GetProperty("recoveryId").GetString());
        Assert.True(recoverArgs.RootElement.GetProperty("show").GetBoolean());
        Assert.Equal(90, recoverArgs.RootElement.GetProperty("timeoutSeconds").GetInt32());
    }

    private static async Task<ServiceRequest> RunAndCaptureAsync(string[] arguments)
    {
        var pipeName = $"excelmcp-safety-cli-{Guid.NewGuid():N}";
        await using var daemon = new CapturingDaemon(pipeName);
        await daemon.StartAsync();

        var result = await CliProcessHelper.RunAsync(
            arguments,
            environmentVariables: new Dictionary<string, string> { ["EXCELMCP_CLI_PIPE"] = pipeName });

        Assert.Equal(0, result.ExitCode);
        return await daemon.CapturedRequest.WaitAsync(TimeSpan.FromSeconds(10));
    }

    private sealed class CapturingDaemon : IAsyncDisposable
    {
        private readonly string _pipeName;
        private readonly Func<ServiceRequest, Task<ServiceResponse>> _handler;
        private readonly CancellationTokenSource _cancellation = new();
        private readonly TaskCompletionSource _ready = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<ServiceRequest> _captured = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly Task _serverTask;
        private NamedPipeServerStream? _currentPipe;

        public CapturingDaemon(string pipeName, Func<ServiceRequest, Task<ServiceResponse>>? handler = null)
        {
            _pipeName = pipeName;
            _handler = handler ?? (request => Task.FromResult(new ServiceResponse
            {
                Success = true,
                Command = request.Command,
                SessionId = request.SessionId,
                Result = "{\"success\":true}"
            }));
            _serverTask = RunAsync();
        }

        public Task<ServiceRequest> CapturedRequest => _captured.Task;

        public Task StartAsync() => _ready.Task;

        public async ValueTask DisposeAsync()
        {
            await _cancellation.CancelAsync();
            if (_currentPipe is not null)
            {
                await _currentPipe.DisposeAsync();
            }

            try
            {
                await _serverTask;
            }
            catch (OperationCanceledException)
            {
            }

            _cancellation.Dispose();
        }

        private async Task RunAsync()
        {
            while (!_cancellation.IsCancellationRequested)
            {
                _currentPipe = ServiceSecurity.CreateSecureServer(_pipeName);
                _ready.TrySetResult();
                await _currentPipe.WaitForConnectionAsync(_cancellation.Token);

                using var rpc = JsonRpc.Attach(_currentPipe, new CapturingTarget(_captured, _handler));
                await rpc.Completion.WaitAsync(_cancellation.Token);
                await _currentPipe.DisposeAsync();
                _currentPipe = null;
            }
        }
    }

    private sealed class CapturingTarget(
        TaskCompletionSource<ServiceRequest> captured,
        Func<ServiceRequest, Task<ServiceResponse>> handler)
    {
        public async Task<ServiceResponse> ProcessCommandAsync(ServiceRequest request)
        {
            if (!request.Command.Equals("service.ping", StringComparison.Ordinal))
            {
                captured.TrySetResult(request);
            }

            return await handler(request);
        }
    }
}
