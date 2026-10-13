using System.Collections.Concurrent;
using System.Globalization;
using System.IO.Compression;
using System.Net;
using System.Net.Sockets;
using System.Text;

namespace Sbroenne.ExcelMcp.Tests.Telemetry;

internal sealed class LocalOtlpReceiver : IAsyncDisposable
{
    private readonly TcpListener _listener = new(IPAddress.Loopback, 0);
    private readonly CancellationTokenSource _shutdown = new();
    private readonly Task _server;
    private readonly byte[] _response;
    private readonly ConcurrentQueue<(string Path, string Payload, IReadOnlyDictionary<string, string> Headers)> _requests = new();

    internal LocalOtlpReceiver(string? redirect = null, string? responseBody = null)
    {
        var responsePayload = responseBody ?? string.Empty;
        var contentType = responseBody is null ? "application/x-protobuf" : "application/json";
        _response = Encoding.UTF8.GetBytes(redirect == null
            ? $"HTTP/1.1 200 OK\r\nContent-Type: {contentType}\r\nContent-Length: {Encoding.UTF8.GetByteCount(responsePayload)}\r\nConnection: close\r\n\r\n{responsePayload}"
            : $"HTTP/1.1 302 Found\r\nLocation: {redirect}\r\nContent-Length: 0\r\nConnection: close\r\n\r\n");
        _listener.Start();
        Endpoint = $"http://127.0.0.1:{((IPEndPoint)_listener.LocalEndpoint).Port}";
        _server = ServeAsync();
    }

    internal string Endpoint { get; }
    internal IReadOnlyList<(string Path, string Payload, IReadOnlyDictionary<string, string> Headers)> Requests => _requests.ToArray();

    private async Task ServeAsync()
    {
        try
        {
            while (true)
            {
                using var client = await _listener.AcceptTcpClientAsync(_shutdown.Token);
                await using var stream = client.GetStream();
                var headerBytes = new List<byte>();
                var oneByte = new byte[1];
                while (true)
                {
                    await stream.ReadExactlyAsync(oneByte, _shutdown.Token);
                    headerBytes.Add(oneByte[0]);
                    if (headerBytes.Count > 16384)
                        throw new InvalidOperationException("Local OTLP request header exceeds the test limit.");
                    if (headerBytes.Count >= 4 && headerBytes.TakeLast(4).SequenceEqual("\r\n\r\n"u8.ToArray()))
                        break;
                }
                var headers = Encoding.ASCII.GetString(headerBytes.ToArray()).Split("\r\n", StringSplitOptions.RemoveEmptyEntries);
                var path = headers[0].Split(' ')[1];
                var contentLength = headers.Single(header => header.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase));
                var size = int.Parse(contentLength.Split(':')[1].Trim(), CultureInfo.InvariantCulture);
                if (size is < 0 or > 1048576)
                    throw new InvalidOperationException("Local OTLP request body exceeds the test limit.");
                var body = new byte[size];
                await stream.ReadExactlyAsync(body, _shutdown.Token);
                var requestHeaders = headers.Skip(1).ToDictionary(
                    header => header[..header.IndexOf(':')],
                    header => header[(header.IndexOf(':') + 1)..].Trim(),
                    StringComparer.OrdinalIgnoreCase);
                if (requestHeaders.TryGetValue("Content-Encoding", out var encoding) &&
                    string.Equals(encoding, "gzip", StringComparison.OrdinalIgnoreCase))
                {
                    using var compressed = new MemoryStream(body);
                    using var gzip = new GZipStream(compressed, CompressionMode.Decompress);
                    using var reader = new StreamReader(gzip, Encoding.UTF8);
                    _requests.Enqueue((path, await reader.ReadToEndAsync(_shutdown.Token), requestHeaders));
                }
                else
                {
                    _requests.Enqueue((path, Encoding.UTF8.GetString(body), requestHeaders));
                }
                await stream.WriteAsync(_response, _shutdown.Token);
            }
        }
        catch (OperationCanceledException) when (_shutdown.IsCancellationRequested)
        {
        }
    }

    public async ValueTask DisposeAsync()
    {
        try
        {
            await _shutdown.CancelAsync();
            await _server.WaitAsync(TimeSpan.FromSeconds(10));
        }
        finally
        {
            _listener.Stop();
            _shutdown.Dispose();
        }
    }
}
