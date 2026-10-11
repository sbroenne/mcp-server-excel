using System.Collections.Concurrent;
using System.Globalization;
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
    private readonly ConcurrentQueue<(string Path, string Payload)> _requests = new();

    internal LocalOtlpReceiver(string? redirect = null)
    {
        _response = Encoding.ASCII.GetBytes(redirect == null
            ? "HTTP/1.1 200 OK\r\nContent-Type: application/x-protobuf\r\nContent-Length: 0\r\nConnection: close\r\n\r\n"
            : $"HTTP/1.1 302 Found\r\nLocation: {redirect}\r\nContent-Length: 0\r\nConnection: close\r\n\r\n");
        _listener.Start();
        Endpoint = $"http://127.0.0.1:{((IPEndPoint)_listener.LocalEndpoint).Port}";
        _server = ServeAsync();
    }

    internal string Endpoint { get; }
    internal IReadOnlyList<(string Path, string Payload)> Requests => _requests.ToArray();

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
                _requests.Enqueue((path, Encoding.UTF8.GetString(body)));
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
