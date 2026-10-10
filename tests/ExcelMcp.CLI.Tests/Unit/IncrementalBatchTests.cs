using System.Text.Json;
using System.Threading.Channels;
using Sbroenne.ExcelMcp.CLI.Commands;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Collection("Service")]
[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "Batch")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class IncrementalBatchTests
{
    [Theory]
    [InlineData("{")]
    [InlineData("null")]
    [InlineData("{}")]
    [InlineData("{\"command\":\" \"}")]
    [InlineData("[]")]
    [InlineData("{\"command\":\"diag.ping\",\"extra\":true}")]
    [InlineData("{\"Command\":\"diag.ping\"}")]
    [InlineData("{\"command\":\"diag.echo\"}")]
    public async Task Stream_InvalidLine_ReturnsErrorBeforeEof_ThenContinues(string invalidLine)
    {
        using var input = new LiveInput();
        using var output = new FlushedOutput();
        using var error = new StringWriter();
        var factory = new RecordingFactory(_ => new ServiceResponse { Success = true, Result = "{\"marker\":123}" });
        var runtime = new CliCommandRuntime(factory, input, output, error, isOutputRedirected: true);
        var invocation = Program.RunAsync(["--quiet", "batch", "--stream"], runtime);
        try
        {
            input.Send(invalidLine);
            using var failed = JsonDocument.Parse(await output.NextLineAsync());
            Assert.False(failed.RootElement.GetProperty("success").GetBoolean());
            Assert.False(string.IsNullOrWhiteSpace(failed.RootElement.GetProperty("error").GetString()));
            Assert.Equal(0, factory.ConnectionCount);
            Assert.False(invocation.IsCompleted);

            input.Send("""{"command":"diag.ping"}""");
            using var valid = JsonDocument.Parse(await output.NextLineAsync());
            Assert.True(valid.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal(1, valid.RootElement.GetProperty("index").GetInt32());
            Assert.Equal(123, valid.RootElement.GetProperty("result").GetProperty("marker").GetInt32());
            input.Complete();
            Assert.Equal(1, await invocation.WaitAsync(TimeSpan.FromSeconds(5)));
            Assert.Single(factory.Requests);
            Assert.True(factory.Disposed);
        }
        finally
        {
            input.Complete();
            await invocation.WaitAsync(TimeSpan.FromSeconds(5));
        }
    }

    [Fact]
    public async Task Stream_StopOnError_ExitsWithoutWaitingForEof()
    {
        using var input = new LiveInput();
        using var output = new FlushedOutput();
        using var error = new StringWriter();
        var factory = new RecordingFactory(_ => throw new InvalidOperationException("Unexpected dispatch"));
        var runtime = new CliCommandRuntime(factory, input, output, error, isOutputRedirected: true);
        var invocation = Program.RunAsync(["--quiet", "batch", "--stream", "--stop-on-error"], runtime);
        input.Send("{");
        input.Send("""{"command":"diag.ping"}""");
        try
        {
            using var failed = JsonDocument.Parse(await output.NextLineAsync());
            Assert.False(failed.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal(1, await invocation.WaitAsync(TimeSpan.FromSeconds(5)));
            Assert.Equal(0, factory.ConnectionCount);
            Assert.Empty(factory.Requests);
        }
        finally { input.Complete(); }
    }

    [Theory]
    [InlineData(false, 2)]
    [InlineData(true, 1)]
    public async Task Stream_NegativeOperationResult_PreservesFailureAndStopPolicy(bool stopOnError, int count)
    {
        var args = new List<string> { "batch", "--stream" };
        if (stopOnError) args.Add("--stop-on-error");
        int calls = 0;
        var result = await InProcessCliHelper.RunAsync(args, _ =>
            new ServiceResponse
            {
                Success = true,
                Result = ++calls == 1 ? "{\"success\":false,\"errorMessage\":\"Occupied cell\"}" : "{\"marker\":456}"
            }, "{\"command\":\"diag.ping\"}\n{\"command\":\"diag.ping\"}\n");
        Assert.Equal(1, result.ExitCode);
        var lines = result.Stdout.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(count, lines.Length);
        Assert.Equal(count, calls);
        using var failed = JsonDocument.Parse(lines[0]);
        Assert.False(failed.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("Occupied cell", failed.RootElement.GetProperty("error").GetString());
        Assert.False(failed.RootElement.GetProperty("result").GetProperty("success").GetBoolean());
    }

    [Theory]
    [InlineData("")]
    [InlineData(" \n\t\n")]
    public async Task Stream_EmptyInput_DoesNotConnect(string input)
    {
        var result = await InProcessCliHelper.RunAsync(["batch", "--stream"], input: input);
        Assert.Equal(1, result.ExitCode);
        Assert.Empty(result.Stdout);
        Assert.Contains("No commands provided", result.Stderr, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Stream_InputFile_IsRejectedWithoutDispatch()
    {
        var result = await InProcessCliHelper.RunAsync(["batch", "--stream", "--input", "commands.json"]);
        Assert.Equal(1, result.ExitCode);
        Assert.Contains("--stream reads NDJSON from stdin", result.Stderr, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Stream_CapturesSession_AllowsOverride_AndClearsOnClose()
    {
        var requests = new List<ServiceRequest>();
        var path = Path.Combine(Path.GetTempPath(), "incremental-session-example.xlsx");
        var input = string.Join('\n',
            " ",
            JsonSerializer.Serialize(new { command = "session.create", args = new { filePath = path } }),
            "{\"command\":\"diag.ping\"}",
            "{\"command\":\"diag.ping\",\"sessionId\":\"override\"}",
            "{\"command\":\"session.close\"}",
            "{\"command\":\"diag.ping\"}");
        var result = await InProcessCliHelper.RunAsync(["batch", "--stream", "--input", "-"], request =>
        {
            requests.Add(request);
            return new ServiceResponse { Success = true, Result = request.Command == "session.create" ? "{\"sessionId\":\"captured\"}" : "{}" };
        }, input);
        Assert.Equal(0, result.ExitCode);
        Assert.Equal(5, requests.Count);
        Assert.Null(requests[0].SessionId);
        Assert.Equal("captured", requests[1].SessionId);
        Assert.Equal("override", requests[2].SessionId);
        Assert.Equal("captured", requests[3].SessionId);
        Assert.Null(requests[4].SessionId);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FailedClose_KeepsSessionForInspection(bool stream)
    {
        var requests = new List<ServiceRequest>();
        string[] args = stream
            ? ["batch", "--stream", "--session", "captured"]
            : ["batch", "--session", "captured"];
        var result = await InProcessCliHelper.RunAsync(args, request =>
        {
            requests.Add(request);
            return new ServiceResponse
            {
                Success = true,
                Result = request.Command == "session.close"
                    ? "{\"success\":false,\"errorMessage\":\"Workbook is busy\"}"
                    : "{\"marker\":123}"
            };
        }, "{\"command\":\"session.close\"}\n{\"command\":\"diag.ping\"}\n");
        Assert.Equal(1, result.ExitCode);
        Assert.Equal(2, requests.Count);
        Assert.Equal("captured", requests[0].SessionId);
        Assert.Equal("captured", requests[1].SessionId);
    }

    [Fact]
    public async Task Stream_InitialConnectionFailure_ReportsIndexedResult_ThenRetries()
    {
        using var input = new LiveInput();
        using var output = new FlushedOutput();
        using var error = new StringWriter();
        var factory = new RecordingFactory(_ => new ServiceResponse { Success = true, Result = "{\"marker\":456}" });
        factory.ConnectFailures.Enqueue(new IOException("Daemon did not become ready"));
        var runtime = new CliCommandRuntime(factory, input, output, error, isOutputRedirected: true);
        var invocation = Program.RunAsync(["--quiet", "batch", "--stream"], runtime);
        try
        {
            input.Send("{\"command\":\"diag.ping\"}");
            using var failed = JsonDocument.Parse(await output.NextLineAsync());
            Assert.Equal(0, failed.RootElement.GetProperty("index").GetInt32());
            Assert.Equal("diag.ping", failed.RootElement.GetProperty("command").GetString());
            Assert.False(failed.RootElement.GetProperty("success").GetBoolean());
            Assert.Contains("Daemon did not become ready", failed.RootElement.GetProperty("error").GetString(), StringComparison.Ordinal);
            Assert.False(invocation.IsCompleted);
            Assert.Empty(factory.Requests);

            input.Send("{\"command\":\"diag.ping\"}");
            using var retried = JsonDocument.Parse(await output.NextLineAsync());
            Assert.Equal(1, retried.RootElement.GetProperty("index").GetInt32());
            Assert.True(retried.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal(456, retried.RootElement.GetProperty("result").GetProperty("marker").GetInt32());
            Assert.Equal(2, factory.ConnectionCount);
            Assert.Single(factory.Requests);

            input.Complete();
            Assert.Equal(1, await invocation.WaitAsync(TimeSpan.FromSeconds(5)));
            Assert.True(factory.Disposed);
        }
        finally
        {
            input.Complete();
            await invocation.WaitAsync(TimeSpan.FromSeconds(5));
        }
    }

    [Fact]
    public async Task Stream_InitialConnectionFailure_WithStopOnError_ReportsResultAndStops()
    {
        using var input = new StringReader("{\"command\":\"diag.ping\"}\n{\"command\":\"diag.ping\"}\n");
        using var output = new StringWriter();
        using var error = new StringWriter();
        var factory = new RecordingFactory(_ => new ServiceResponse { Success = true });
        factory.ConnectFailures.Enqueue(new IOException("Daemon did not become ready"));
        var runtime = new CliCommandRuntime(factory, input, output, error, isOutputRedirected: true);

        Assert.Equal(1, await Program.RunAsync(["--quiet", "batch", "--stream", "--stop-on-error"], runtime));

        var line = Assert.Single(output.ToString().Split('\n', StringSplitOptions.RemoveEmptyEntries));
        using var failed = JsonDocument.Parse(line);
        Assert.Equal(0, failed.RootElement.GetProperty("index").GetInt32());
        Assert.False(failed.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(1, factory.ConnectionCount);
        Assert.Empty(factory.Requests);
    }

    [Fact]
    public async Task Stream_CancellationDuringConnect_Propagates()
    {
        using var input = new StringReader("{\"command\":\"diag.ping\"}\n");
        using var output = new StringWriter();
        using var error = new StringWriter();
        using var cancellation = new CancellationTokenSource();
        var factory = new RecordingFactory(_ => new ServiceResponse { Success = true });
        factory.OnConnect = () =>
        {
            cancellation.Cancel();
            cancellation.Token.ThrowIfCancellationRequested();
        };
        using var scope = CliCommandRuntime.Push(new CliCommandRuntime(factory, input, output, error, isOutputRedirected: true));

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new BatchCommand().ExecuteAsync(null!, new BatchCommand.Settings { Stream = true }, cancellation.Token));
        Assert.Equal(string.Empty, output.ToString());
        Assert.Empty(factory.Requests);
    }

    [Fact]
    public async Task Stream_CommunicationFailure_IsReported_AndClientIsDisposed()
    {
        using var input = new StringReader("{\"command\":\"diag.ping\"}\n{\"command\":\"diag.ping\"}\n");
        using var output = new StringWriter();
        using var error = new StringWriter();
        var factory = new RecordingFactory(_ => throw new IOException("Connection lost"));
        var runtime = new CliCommandRuntime(factory, input, output, error, isOutputRedirected: true);
        Assert.Equal(1, await Program.RunAsync(["--quiet", "batch", "--stream"], runtime));
        Assert.Equal(2, factory.Requests.Count);
        Assert.True(factory.Disposed);
        foreach (var line in output.ToString().Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            using var result = JsonDocument.Parse(line);
            Assert.False(result.RootElement.GetProperty("success").GetBoolean());
            Assert.Contains("Connection lost", result.RootElement.GetProperty("error").GetString(), StringComparison.Ordinal);
        }
    }

    [Fact]
    public async Task Stream_CancellationWhileWaiting_DisposesClientWithoutSuccess()
    {
        using var input = new LiveInput();
        using var output = new FlushedOutput();
        using var error = new StringWriter();
        using var cancellation = new CancellationTokenSource();
        var factory = new RecordingFactory(_ => new ServiceResponse { Success = true, Result = "{\"marker\":789}" });
        using var scope = CliCommandRuntime.Push(new CliCommandRuntime(factory, input, output, error, isOutputRedirected: true));
        var invocation = new BatchCommand().ExecuteAsync(null!, new BatchCommand.Settings { Stream = true }, cancellation.Token);
        input.Send("{\"command\":\"diag.ping\"}");
        try
        {
            using var first = JsonDocument.Parse(await output.NextLineAsync());
            Assert.Equal(789, first.RootElement.GetProperty("result").GetProperty("marker").GetInt32());
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => invocation.WaitAsync(TimeSpan.FromSeconds(5)));
            Assert.True(factory.Disposed);
            Assert.Single(factory.Requests);
            Assert.Single(output.ToString().Split('\n', StringSplitOptions.RemoveEmptyEntries));
        }
        finally { input.Complete(); }
    }

    [Fact]
    public async Task Stream_ReturnsFlushedResultBeforeEof_ThenAcceptsDependentCommand()
    {
        using var input = new LiveInput();
        using var output = new FlushedOutput();
        using var error = new StringWriter();
        var factory = new RecordingFactory(request =>
        {
            using var args = JsonDocument.Parse(request.Args!);
            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new { message = args.RootElement.GetProperty("message").GetString() })
            };
        });
        var runtime = new CliCommandRuntime(factory, input, output, error, isOutputRedirected: true);
        var invocation = Program.RunAsync(["--quiet", "batch", "--stream"], runtime);
        try
        {
            input.Send("""{"command":"diag.echo","args":{"message":"first"}}""");
            using var first = JsonDocument.Parse(await output.NextLineAsync());
            Assert.True(first.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal(0, first.RootElement.GetProperty("index").GetInt32());
            Assert.False(invocation.IsCompleted);

            var nextMessage = first.RootElement.GetProperty("result").GetProperty("message").GetString() + " then second";
            input.Send(JsonSerializer.Serialize(new { command = "diag.echo", args = new { message = nextMessage } }));
            using var second = JsonDocument.Parse(await output.NextLineAsync());
            Assert.Equal("first then second", second.RootElement.GetProperty("result").GetProperty("message").GetString());
            Assert.Equal(1, second.RootElement.GetProperty("index").GetInt32());
            Assert.False(invocation.IsCompleted);
            Assert.Equal(1, factory.ConnectionCount);
            Assert.Equal(2, factory.Requests.Count);
            input.Complete();
            Assert.Equal(0, await invocation.WaitAsync(TimeSpan.FromSeconds(5)));
            Assert.True(factory.Disposed);
            Assert.Equal(string.Empty, error.ToString());
        }
        finally
        {
            input.Complete();
            await invocation.WaitAsync(TimeSpan.FromSeconds(5));
        }
    }

    private sealed class LiveInput : TextReader
    {
        private readonly Channel<string> _lines = Channel.CreateUnbounded<string>();

        internal void Send(string line) => Assert.True(_lines.Writer.TryWrite(line));
        internal void Complete() => _lines.Writer.TryComplete();

        public override async ValueTask<string?> ReadLineAsync(CancellationToken cancellationToken)
        {
            while (await _lines.Reader.WaitToReadAsync(cancellationToken))
            {
                if (_lines.Reader.TryRead(out var line)) return line;
            }
            return null;
        }

        protected override void Dispose(bool disposing)
        {
            Complete();
            base.Dispose(disposing);
        }
    }

    // Results are observable only on flush, as with a buffered pipe writer.
    private sealed class FlushedOutput : StringWriter
    {
        private readonly Queue<string> _pending = new();
        private readonly Channel<string> _flushed = Channel.CreateUnbounded<string>();

        public override void WriteLine(string? value)
        {
            base.WriteLine(value);
            _pending.Enqueue(value ?? string.Empty);
        }

        public override Task FlushAsync(CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            while (_pending.TryDequeue(out var line)) _flushed.Writer.TryWrite(line);
            return Task.CompletedTask;
        }

        internal async Task<string> NextLineAsync() =>
            await _flushed.Reader.ReadAsync().AsTask().WaitAsync(TimeSpan.FromSeconds(5));
    }

    private sealed class RecordingFactory(Func<ServiceRequest, ServiceResponse> responder) : ICliRequestClientFactory
    {
        private readonly Func<ServiceRequest, ServiceResponse> _responder = responder;
        internal List<ServiceRequest> Requests { get; } = [];
        internal int ConnectionCount { get; private set; }
        internal bool Disposed { get; private set; }
        internal Queue<Exception> ConnectFailures { get; } = new();
        internal Action? OnConnect { get; set; }

        public Task<ICliRequestClient> ConnectAsync(CancellationToken cancellationToken)
        {
            ConnectionCount++;
            OnConnect?.Invoke();
            if (ConnectFailures.TryDequeue(out var failure))
            {
                return Task.FromException<ICliRequestClient>(failure);
            }
            return Task.FromResult<ICliRequestClient>(new Client(this));
        }

        private sealed class Client(RecordingFactory owner) : ICliRequestClient
        {
            public Task<ServiceResponse> SendAsync(ServiceRequest request, CancellationToken cancellationToken)
            {
                owner.Requests.Add(request);
                return Task.FromResult(owner._responder(request));
            }

            public void Dispose() => owner.Disposed = true;
        }
    }
}
