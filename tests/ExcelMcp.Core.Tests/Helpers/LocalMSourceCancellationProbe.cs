using System.Net;
using System.Net.Sockets;
using System.Text;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

internal sealed class LocalMSourceCancellationProbe : IAsyncDisposable
{
    private readonly TcpListener _listener = new(IPAddress.Loopback, 0);
    private readonly CancellationTokenSource _shutdown = new();
    private readonly CancellationTokenSource _operation;
    private readonly Task _server;
    private int _requests;
    private int _cancelledBeforeResponse;

    public LocalMSourceCancellationProbe(CancellationTokenSource operation)
    {
        _operation = operation;
        _listener.Start();
        Url = $"http://127.0.0.1:{((IPEndPoint)_listener.LocalEndpoint).Port}/evaluation.csv";
        _server = ServeAsync();
    }

    public string Url { get; }
    public int Requests => Volatile.Read(ref _requests);
    public bool CancelledBeforeResponse => Volatile.Read(ref _cancelledBeforeResponse) != 0;

    private async Task ServeAsync()
    {
        try
        {
            while (true)
            {
                using var client = await _listener.AcceptTcpClientAsync(_shutdown.Token);
                await using var stream = client.GetStream();
                using var reader = new StreamReader(stream, Encoding.ASCII, leaveOpen: true);
                var request = await reader.ReadLineAsync(_shutdown.Token);
                if (request != "GET /evaluation.csv HTTP/1.1")
                    throw new InvalidOperationException($"Unexpected local M-source request: {request}");
                while (!string.IsNullOrEmpty(await reader.ReadLineAsync(_shutdown.Token)))
                {
                }
                Interlocked.Increment(ref _requests);
                _operation.Cancel();
                Interlocked.Exchange(ref _cancelledBeforeResponse, _operation.IsCancellationRequested ? 1 : 0);
                var response = Encoding.ASCII.GetBytes(
                    "HTTP/1.1 200 OK\r\nContent-Type: text/csv\r\nContent-Length: 10\r\nConnection: close\r\n\r\nValue\r\n1\r\n");
                await stream.WriteAsync(response, _shutdown.Token);
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
