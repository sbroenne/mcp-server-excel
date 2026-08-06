using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;

namespace Sbroenne.ExcelMcp.Benchmarks;

internal sealed record ProtocolFootprintResult(
    IReadOnlyList<BenchmarkObservation> Observations,
    int ToolCount,
    bool ToolCallSucceeded,
    long InitializationPayloadBytes);

internal static class ProtocolFootprintProbe
{
    public static async Task<ProtocolFootprintResult> RunAsync(
        int iterations,
        string testWorkbookPath,
        CancellationToken cancellationToken)
    {
        var clientToServer = new Pipe();
        var serverToClient = new Pipe();
        await using var clientInput = new CaptureStream(clientToServer.Writer.AsStream(), captureWrites: true, leaveOpen: true);
        await using var clientOutput = new CaptureStream(serverToClient.Reader.AsStream(), captureWrites: false, leaveOpen: true);
        McpClient? client = null;
        Task<int>? serverTask = null;

        try
        {
            Sbroenne.ExcelMcp.McpServer.Program.ConfigureTestTransport(clientToServer, serverToClient);
            serverTask = Sbroenne.ExcelMcp.McpServer.Program.Main([]);
            client = await McpClient.CreateAsync(
                new StreamClientTransport(serverInput: clientInput, serverOutput: clientOutput),
                clientOptions: new McpClientOptions
                {
                    ClientInfo = new() { Name = "ExcelMcpBenchmark", Version = "1.0.0" },
                    InitializationTimeout = TimeSpan.FromSeconds(30)
                },
                cancellationToken: cancellationToken);

            var initializationBytes = clientInput.CapturedLength + clientOutput.CapturedLength;
            var observations = new List<BenchmarkObservation>(iterations);
            var toolCount = 0;
            var toolCallSucceeded = true;

            for (var iteration = 0; iteration < iterations; iteration++)
            {
                var requestStart = clientInput.CapturedLength;
                var responseStart = clientOutput.CapturedLength;
                var tools = await client.ListToolsAsync(cancellationToken: cancellationToken);
                toolCount = tools.Count;
                var requestBytes = clientInput.CapturedLength - requestStart;
                var responseBytes = clientOutput.CapturedLength - responseStart;
                var schemaBytes = requestBytes + responseBytes;
                var schemaHash = BenchmarkContext.Sha256(Convert.ToBase64String(clientOutput.GetCapturedBytes(responseStart)));

                requestStart = clientInput.CapturedLength;
                responseStart = clientOutput.CapturedLength;
                var callResult = await client.CallToolAsync(
                    "file",
                    new Dictionary<string, object?>
                    {
                        ["action"] = "test",
                        ["path"] = testWorkbookPath
                    },
                    cancellationToken: cancellationToken);
                var callBytes = (clientInput.CapturedLength - requestStart) + (clientOutput.CapturedLength - responseStart);
                var callSucceeded = IsSuccessfulJsonResponse(callResult);
                toolCallSucceeded &= callSucceeded;

                observations.Add(new BenchmarkObservation(
                    iteration,
                    "mcp-wire-footprint",
                    tools.Count > 0 && callSucceeded,
                    null,
                    new Dictionary<string, double>
                    {
                        ["mcp_schema_payload_bytes"] = schemaBytes,
                        ["mcp_schema_token_estimate"] = BenchmarkContext.EstimateTokensFromUtf8Bytes(schemaBytes),
                        ["schema_response_bytes"] = responseBytes,
                        ["schema_request_bytes"] = requestBytes,
                        ["tool_call_payload_bytes"] = callBytes,
                        ["tool_call_token_estimate"] = BenchmarkContext.EstimateTokensFromUtf8Bytes(callBytes),
                        ["tool_count"] = tools.Count,
                        ["initialization_payload_bytes"] = initializationBytes
                    },
                    new Dictionary<string, string>
                    {
                        ["token_measurement"] = "ceil(utf8-wire-bytes/4); deterministic estimate, not model-specific",
                        ["schema_response_sha256"] = schemaHash
                    },
                    "protocol-success"));
            }

            return new ProtocolFootprintResult(observations, toolCount, toolCallSucceeded, initializationBytes);
        }
        finally
        {
            if (client is not null)
            {
                await client.DisposeAsync();
            }

            Sbroenne.ExcelMcp.McpServer.Program.RequestTestTransportShutdown();
            await TryCompleteAsync(clientToServer.Writer);
            await TryCompleteAsync(serverToClient.Reader);

            if (serverTask is not null)
            {
                try
                {
                    await serverTask.WaitAsync(TimeSpan.FromSeconds(30), cancellationToken);
                }
                catch (OperationCanceledException)
                {
                }
                catch (TimeoutException)
                {
                }
            }

            await TryCompleteAsync(clientToServer.Reader);
            await TryCompleteAsync(serverToClient.Writer);
            Sbroenne.ExcelMcp.McpServer.Program.ResetTestTransport();
        }
    }

    private static async Task TryCompleteAsync(PipeReader reader)
    {
        try
        {
            await reader.CompleteAsync();
        }
        catch (InvalidOperationException)
        {
        }
    }

    private static bool IsSuccessfulJsonResponse(CallToolResult result)
    {
        var text = result.Content.OfType<TextContentBlock>().FirstOrDefault()?.Text;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        try
        {
            using var document = JsonDocument.Parse(text);
            return document.RootElement.TryGetProperty("success", out var success) && success.GetBoolean();
        }
        catch (JsonException)
        {
            return false;
        }
    }

    private static async Task TryCompleteAsync(PipeWriter writer)
    {
        try
        {
            await writer.CompleteAsync();
        }
        catch (InvalidOperationException)
        {
        }
    }
}

internal sealed class CaptureStream : Stream
{
    private readonly Stream _inner;
    private readonly bool _captureWrites;
    private readonly bool _leaveOpen;
    private readonly MemoryStream _capture = new();
    private readonly object _sync = new();

    public CaptureStream(Stream inner, bool captureWrites, bool leaveOpen)
    {
        _inner = inner;
        _captureWrites = captureWrites;
        _leaveOpen = leaveOpen;
    }

    public long CapturedLength
    {
        get
        {
            lock (_sync)
            {
                return _capture.Length;
            }
        }
    }

    public override bool CanRead => _inner.CanRead;

    public override bool CanSeek => false;

    public override bool CanWrite => _inner.CanWrite;

    public override long Length => throw new NotSupportedException();

    public override long Position
    {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    public byte[] GetCapturedBytes(long start)
    {
        lock (_sync)
        {
            var bytes = _capture.ToArray();
            if (start < 0 || start > bytes.LongLength)
            {
                throw new ArgumentOutOfRangeException(nameof(start));
            }

            return bytes[(int)start..];
        }
    }

    public override void Flush() => _inner.Flush();

    public override Task FlushAsync(CancellationToken cancellationToken) => _inner.FlushAsync(cancellationToken);

    public override int Read(byte[] buffer, int offset, int count)
    {
        var read = _inner.Read(buffer, offset, count);
        if (!_captureWrites && read > 0)
        {
            Capture(buffer.AsSpan(offset, read));
        }

        return read;
    }

    public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
    {
        var read = await _inner.ReadAsync(buffer, cancellationToken);
        if (!_captureWrites && read > 0)
        {
            Capture(buffer.Span[..read]);
        }

        return read;
    }

    public override void Write(byte[] buffer, int offset, int count)
    {
        if (_captureWrites && count > 0)
        {
            Capture(buffer.AsSpan(offset, count));
        }

        _inner.Write(buffer, offset, count);
    }

    public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
    {
        if (_captureWrites && buffer.Length > 0)
        {
            Capture(buffer.Span);
        }

        await _inner.WriteAsync(buffer, cancellationToken);
    }

    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

    public override void SetLength(long value) => throw new NotSupportedException();

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _capture.Dispose();
            if (!_leaveOpen)
            {
                _inner.Dispose();
            }
        }

        base.Dispose(disposing);
    }

    private void Capture(ReadOnlySpan<byte> buffer)
    {
        lock (_sync)
        {
            _capture.Write(buffer);
        }
    }
}
