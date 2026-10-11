using System.Buffers.Binary;
using System.Collections.Concurrent;
using System.Net;
using System.Text;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Server.Kestrel.Core;
using Microsoft.Extensions.Logging;

namespace Sbroenne.ExcelMcp.Tests.Telemetry;

internal sealed class LocalGrpcOtlpReceiver : IAsyncDisposable
{
    private readonly WebApplication _server;
    private readonly ConcurrentQueue<(string Path, string Payload, IReadOnlyDictionary<string, string> Headers)> _requests = new();

    private LocalGrpcOtlpReceiver(WebApplication server) => _server = server;

    internal string Endpoint => _server.Urls.Single();
    internal IReadOnlyList<(string Path, string Payload, IReadOnlyDictionary<string, string> Headers)> Requests => _requests.ToArray();

    internal static async Task<LocalGrpcOtlpReceiver> StartAsync()
    {
        var builder = WebApplication.CreateSlimBuilder();
        builder.Logging.ClearProviders();
        builder.WebHost.ConfigureKestrel(options =>
            options.Listen(IPAddress.Loopback, 0, listen => listen.Protocols = HttpProtocols.Http2));
        var server = builder.Build();
        var receiver = new LocalGrpcOtlpReceiver(server);
        server.Run(receiver.ReceiveAsync);
        try
        {
            await server.StartAsync();
            return receiver;
        }
        catch
        {
            await server.DisposeAsync();
            throw;
        }
    }

    private async Task ReceiveAsync(HttpContext context)
    {
        if (context.Request.Protocol != "HTTP/2" || context.Request.Method != "POST"
            || context.Request.ContentType != "application/grpc")
            throw new InvalidOperationException("Expected an HTTP/2 gRPC export request.");

        var frame = new byte[5];
        await context.Request.Body.ReadExactlyAsync(frame, context.RequestAborted);
        var size = BinaryPrimitives.ReadInt32BigEndian(frame.AsSpan(1));
        if (frame[0] != 0 || size is < 0 or > 1048576)
            throw new InvalidOperationException("Invalid or oversized gRPC message frame.");
        var body = new byte[size];
        await context.Request.Body.ReadExactlyAsync(body, context.RequestAborted);
        var trailing = new byte[1];
        if (await context.Request.Body.ReadAsync(trailing, context.RequestAborted) != 0)
            throw new InvalidOperationException("Expected exactly one gRPC export message.");
        var headers = context.Request.Headers.ToDictionary(
            header => header.Key, header => header.Value.ToString(), StringComparer.OrdinalIgnoreCase);
        _requests.Enqueue((context.Request.Path.Value!, Encoding.UTF8.GetString(body), headers));

        // An empty protobuf ExportResponse still needs a gRPC frame and success trailer.
        context.Response.ContentType = "application/grpc";
        context.Response.DeclareTrailer("grpc-status");
        await context.Response.Body.WriteAsync(new byte[5], context.RequestAborted);
        context.Response.AppendTrailer("grpc-status", "0");
    }

    public async ValueTask DisposeAsync()
    {
        try
        {
            using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(10));
            await _server.StopAsync(deadline.Token);
        }
        finally
        {
            await _server.DisposeAsync();
        }
    }
}
