using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Integration tests for the batch CLI command.
/// Tests real CLI parsing, Service dispatch, and NDJSON output without Excel.
/// The in-process tests keep the commit hook independent of child processes.
/// </summary>
[Collection("Service")]
[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "Batch")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class BatchCommandTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;

    public BatchCommandTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Combine(Path.GetTempPath(), $"BatchTest_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        try { Directory.Delete(_tempDir, recursive: true); }
        catch { /* best effort */ }
    }

    [Theory]
    [InlineData("create")]
    [InlineData("update")]
    public async Task Batch_PowerQueryFailure_PreservesDiagnosticsAndPartialState(string action)
    {
        var partialState = JsonSerializer.SerializeToElement(new
        {
            queryName = "Query",
            queryExists = true,
            loadMode = "load-to-table",
            targetSheet = "Actual",
            isLoadedToDataModel = false
        });
        var result = await InProcessCliHelper.RunAsync(["batch", "--session", "session"],
            request =>
            {
                Assert.Equal($"powerquery.{action}", request.Command);
                Assert.Equal("session", request.SessionId);
                return new ServiceResponse
                {
                    Success = false,
                    ErrorMessage = "Query remains after privacy failure",
                    ErrorCategory = "Privacy",
                    ExceptionType = "PowerQueryCommandException",
                    HResult = "0x800A03EC",
                    InnerError = "Original privacy error",
                    PartialState = partialState.GetRawText()
                };
            },
            $$$"""{"command":"powerquery.{{{action}}}","args":{"queryName":"Query","mCode":"error \"Formula.Firewall\""}}""");

        Assert.Equal(1, result.ExitCode);
        Assert.Equal(
            $"Batch 1/1: powerquery.{action} started.{Environment.NewLine}" +
            $"Batch 1/1: powerquery.{action} failed.{Environment.NewLine}",
            result.Stderr);
        using var json = JsonDocument.Parse(result.Stdout);
        var root = json.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean());
        Assert.Equal(partialState.GetRawText(), root.GetProperty("partialState").GetRawText());
        Assert.Equal("Privacy", root.GetProperty("error").GetProperty("category").GetString());
        Assert.Equal("PowerQueryCommandException", root.GetProperty("error").GetProperty("exceptionType").GetString());
        Assert.Equal("0x800A03EC", root.GetProperty("error").GetProperty("hresult").GetString());
        Assert.Equal("Original privacy error", root.GetProperty("error").GetProperty("innerError").GetString());
        Assert.Equal("Query remains after privacy failure", root.GetProperty("error").GetProperty("message").GetString());
    }

    // ============================================
    // SINGLE COMMAND - Basic execution
    // ============================================

    // ============================================
    // MULTIPLE COMMANDS - Sequential execution
    // ============================================

    [Fact]
    public async Task Batch_BlockedLaterCommand_AlreadyFlushedEarlierResultAndCurrentProgress()
    {
        using var output = new FlushRecordingWriter();
        using var error = new FlushRecordingWriter();
        var client = new DeferredBatchClient();
        var runtime = new CliCommandRuntime(client,
            new StringReader("""[{"command":"diag.ping"},{"command":"diag.echo","args":{"message":"later"}}]"""),
            output, error, isOutputRedirected: true);
        var execution = Program.RunAsync(["--quiet", "batch"], runtime);
        try
        {
            await client.SecondStarted.Task.WaitAsync(TimeSpan.FromSeconds(5));
            Assert.False(execution.IsCompleted);
            Assert.Equal(1, output.Flushes);
            Assert.Equal(2, error.Flushes);
            using var first = JsonDocument.Parse(output.ToString());
            Assert.True(first.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal(0, first.RootElement.GetProperty("index").GetInt32());
            Assert.Equal(
                $"Batch 1/2: diag.ping started.{Environment.NewLine}" +
                $"Batch 1/2: diag.ping completed.{Environment.NewLine}" +
                $"Batch 2/2: diag.echo started.{Environment.NewLine}", error.ToString());
        }
        finally
        {
            client.Finish.TrySetResult(new ServiceResponse { Success = true, Result = """{"message":"later"}""" });
            Assert.Equal(0, await execution.WaitAsync(TimeSpan.FromSeconds(5)));
        }
        Assert.Equal(2, output.Flushes);
        var lines = output.ToString().Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(2, lines.Length);
        using var last = JsonDocument.Parse(lines[1]);
        Assert.Equal(1, last.RootElement.GetProperty("index").GetInt32());
        Assert.Equal("later", last.RootElement.GetProperty("data").GetProperty("message").GetString());
    }

    private sealed class FlushRecordingWriter : StringWriter
    {
        internal int Flushes { get; private set; }

        public override Task FlushAsync(CancellationToken cancellationToken)
        {
            Flushes++;
            return base.FlushAsync(cancellationToken);
        }
    }

    private sealed class DeferredBatchClient : ICliRequestClientFactory, ICliRequestClient
    {
        internal TaskCompletionSource SecondStarted { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource<ServiceResponse> Finish { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _calls;

        public Task<ICliRequestClient> ConnectAsync(CancellationToken cancellationToken) =>
            Task.FromResult<ICliRequestClient>(this);

        public Task<ServiceResponse> SendAsync(ServiceRequest request, CancellationToken cancellationToken)
        {
            if (++_calls == 1)
                return Task.FromResult(new ServiceResponse { Success = true, Result = """{"message":"pong"}""" });
            SecondStarted.SetResult();
            return Finish.Task;
        }

        public void Dispose() { }
    }

    [Fact]
    public async Task Batch_MultipleCommands_ExecutesAllSequentially()
    {
        var inputFile = WriteBatchFile("""
            [
              {"command": "diag.ping"},
              {"command": "diag.echo", "args": {"message": "hello"}},
              {"command": "diag.echo", "args": {"message": "world"}}
            ]
            """);

        var results = await RunBatchAsync($"batch --input \"{inputFile}\"");

        Assert.Equal(3, results.Count);

        // All should succeed
        Assert.All(results, r => Assert.True(r.GetProperty("success").GetBoolean()));

        // Indices should be sequential
        Assert.Equal(0, results[0].GetProperty("index").GetInt32());
        Assert.Equal(1, results[1].GetProperty("index").GetInt32());
        Assert.Equal(2, results[2].GetProperty("index").GetInt32());

        // Verify echo results
        Assert.Equal("hello", results[1].GetProperty("data").GetProperty("message").GetString());
        Assert.Equal("world", results[2].GetProperty("data").GetProperty("message").GetString());
    }

    // ============================================
    // ERROR HANDLING - Continue on error (default)
    // ============================================

    [Fact]
    public async Task Batch_CommandFails_ContinuesRemainingByDefault()
    {
        var inputFile = WriteBatchFile("""
            [
              {"command": "diag.ping"},
              {"command": "nonexistent.action"},
              {"command": "diag.ping"}
            ]
            """);

        var results = await RunBatchAsync($"batch --input \"{inputFile}\"");

        Assert.Equal(3, results.Count);
        Assert.True(results[0].GetProperty("success").GetBoolean());
        Assert.False(results[1].GetProperty("success").GetBoolean()); // invalid command fails
        Assert.True(results[2].GetProperty("success").GetBoolean()); // should still execute
    }

    // ============================================
    // ERROR HANDLING - Stop on error
    // ============================================

    [Fact]
    public async Task Batch_StopOnError_StopsAfterFirstFailure()
    {
        var inputFile = WriteBatchFile("""
            [
              {"command": "diag.ping"},
              {"command": "nonexistent.action"},
              {"command": "diag.ping"}
            ]
            """);

        var results = await RunBatchAsync($"batch --input \"{inputFile}\" --stop-on-error");

        Assert.Equal(2, results.Count); // Only 2 results — stopped after failure
        Assert.True(results[0].GetProperty("success").GetBoolean());
        Assert.False(results[1].GetProperty("success").GetBoolean());
    }

    // ============================================
    // INPUT VALIDATION
    // ============================================

    [Fact]
    public async Task Batch_EmptyArray_ReturnsError()
    {
        var inputFile = WriteBatchFile("[]");

        var result = await InProcessCliHelper.RunWithServiceAsync($"batch --input \"{inputFile}\"");
        _output.WriteLine($"Exit: {result.ExitCode}, Stderr: {result.Stderr}");

        Assert.Equal(1, result.ExitCode);
    }

    [Fact]
    public async Task Batch_MissingCommandField_ReturnsError()
    {
        var inputFile = WriteBatchFile("""
            [
              {"args": {"message": "no command field"}}
            ]
            """);

        var result = await InProcessCliHelper.RunWithServiceAsync($"batch --input \"{inputFile}\"");
        _output.WriteLine($"Exit: {result.ExitCode}, Stdout: {result.Stdout}");

        Assert.Equal(1, result.ExitCode);
    }

    [Fact]
    public async Task Batch_FileNotFound_ReturnsError()
    {
        var result = await InProcessCliHelper.RunWithServiceAsync("batch --input \"nonexistent.json\"");
        _output.WriteLine($"Exit: {result.ExitCode}, Stderr: {result.Stderr}");

        Assert.Equal(1, result.ExitCode);
    }

    // ============================================
    // NDJSON FORMAT - Output structure
    // ============================================

    [Fact]
    public async Task Batch_OutputIsValidNdjson_OneLinePerCommand()
    {
        var inputFile = WriteBatchFile("""
            [
              {"command": "diag.ping"},
              {"command": "diag.echo", "args": {"message": "test"}}
            ]
            """);

        var result = await InProcessCliHelper.RunWithServiceAsync($"batch --input \"{inputFile}\"");
        _output.WriteLine(result.Stdout);

        Assert.Equal(0, result.ExitCode);

        // Each non-empty line should be valid JSON
        var lines = result.Stdout.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(2, lines.Length);

        foreach (var line in lines)
        {
            var json = JsonDocument.Parse(line.Trim());
            Assert.True(json.RootElement.TryGetProperty("index", out _));
            Assert.True(json.RootElement.TryGetProperty("command", out _));
            Assert.True(json.RootElement.TryGetProperty("success", out _));
        }
    }

    [Fact]
    public async Task Batch_SuccessResult_HasNoErrorMessageField()
    {
        var inputFile = WriteBatchFile("""
            [
              {"command": "diag.ping"}
            ]
            """);

        var results = await RunBatchAsync($"batch --input \"{inputFile}\"");

        Assert.Single(results);
        Assert.True(results[0].GetProperty("success").GetBoolean());
        // Error message should be absent (JsonIgnoreCondition.WhenWritingNull).
        Assert.False(results[0].TryGetProperty("errorMessage", out _),
            "Success result should not contain 'errorMessage' field");
        Assert.False(results[0].TryGetProperty("partialState", out _));
        Assert.False(results[0].TryGetProperty("errorCategory", out _));
    }

    [Fact]
    public async Task Batch_ErrorResult_HasStructuredErrorWithoutOldMessageField()
    {
        var inputFile = WriteBatchFile("""
            [
              {"command": "nonexistent.action"}
            ]
            """);

        var results = await RunBatchAsync($"batch --input \"{inputFile}\"");

        Assert.Single(results);
        Assert.False(results[0].GetProperty("success").GetBoolean());
        Assert.True(results[0].GetProperty("error").TryGetProperty("message", out var errorMessage));
        Assert.False(string.IsNullOrWhiteSpace(errorMessage.GetString()));
        Assert.False(results[0].TryGetProperty("errorMessage", out _));
        Assert.False(results[0].TryGetProperty("partialState", out _));
    }

    // ============================================
    // NDJSON INPUT FORMAT
    // ============================================

    [Fact]
    public async Task Batch_NdjsonInputFormat_ParsesCorrectly()
    {
        // NDJSON: one command per line (no array wrapper)
        var inputFile = WriteBatchFile(
            """{"command": "diag.ping"}""" + "\n" +
            """{"command": "diag.echo", "args": {"message": "ndjson"}}""");

        var results = await RunBatchAsync($"batch --input \"{inputFile}\"");

        Assert.Equal(2, results.Count);
        Assert.True(results[0].GetProperty("success").GetBoolean());
        Assert.Equal("ndjson", results[1].GetProperty("data").GetProperty("message").GetString());
    }

    // ============================================
    // EXIT CODE
    // ============================================

    [Fact]
    public async Task Batch_AllSucceed_ExitCodeZero()
    {
        var inputFile = WriteBatchFile("""
            [
              {"command": "diag.ping"},
              {"command": "diag.echo", "args": {"message": "test"}}
            ]
            """);

        var result = await InProcessCliHelper.RunWithServiceAsync($"batch --input \"{inputFile}\"");
        Assert.Equal(0, result.ExitCode);
    }

    [Fact]
    public async Task Batch_AnyFails_ExitCodeOne()
    {
        var inputFile = WriteBatchFile("""
            [
              {"command": "diag.ping"},
              {"command": "nonexistent.action"}
            ]
            """);

        var result = await InProcessCliHelper.RunWithServiceAsync($"batch --input \"{inputFile}\"");
        Assert.Equal(1, result.ExitCode);
    }

    // ============================================
    // Helpers
    // ============================================

    private string WriteBatchFile(string content)
    {
        var path = Path.Combine(_tempDir, $"batch_{Guid.NewGuid():N}.json");
        File.WriteAllText(path, content);
        return path;
    }

    /// <summary>
    /// Runs a batch command and parses the NDJSON output into a list of JsonElements.
    /// </summary>
    private async Task<List<JsonElement>> RunBatchAsync(
        string args)
    {
        var result = await InProcessCliHelper.RunWithServiceAsync(args);
        _output.WriteLine($"Exit: {result.ExitCode}");
        _output.WriteLine($"Stdout: {result.Stdout}");
        if (!string.IsNullOrEmpty(result.Stderr))
            _output.WriteLine($"Stderr: {result.Stderr}");

        var results = new List<JsonElement>();
        foreach (var line in result.Stdout.Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            var trimmed = line.Trim();
            if (string.IsNullOrEmpty(trimmed)) continue;

            try
            {
                using var doc = JsonDocument.Parse(trimmed);
                results.Add(doc.RootElement.Clone());
            }
            catch (JsonException)
            {
                _output.WriteLine($"Non-JSON line: {trimmed}");
            }
        }

        return results;
    }
}

[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "Batch")]
[Trait("RequiresExcel", "false")]
[Trait("AdapterTestKind", "System")]
[Trait("Speed", "Medium")]
[Collection("Service")]
public sealed class BatchProcessSmokeTests
{
    [Fact]
    public async Task Batch_SinglePingCommand_ReturnsSuccess()
    {
        var inputFile = Path.Combine(Path.GetTempPath(), $"BatchProcessSmoke_{Guid.NewGuid():N}.json");
        await File.WriteAllTextAsync(inputFile, """
            [
              {"command": "diag.ping"}
            ]
            """);

        try
        {
            var result = await CliProcessHelper.RunAsync($"batch --input \"{inputFile}\"");
            Assert.Equal(0, result.ExitCode);
            using var json = JsonDocument.Parse(result.Stdout.Trim());
            Assert.Equal(0, json.RootElement.GetProperty("index").GetInt32());
            Assert.Equal("diag.ping", json.RootElement.GetProperty("command").GetString());
            Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        }
        finally
        {
            File.Delete(inputFile);
        }
    }
}
