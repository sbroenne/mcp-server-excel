using System.IO.Compression;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.Json.Nodes;

namespace Sbroenne.ExcelMcp.CLI.Tests.Helpers;

/// <summary>
/// In-process stand-in for the Application Insights ingestion endpoint. It listens on
/// a free loopback port, records every envelope posted to <c>/v2.1/track</c> and
/// acknowledges it, so tests can inspect exactly what a telemetry sink would send.
/// Nothing is forwarded anywhere.
/// </summary>
internal sealed class FakeIngestionEndpoint : IDisposable
{
    internal const string InstrumentationKey = "00000000-0000-0000-0000-000000000001";

    private readonly HttpListener _listener = new();
    private readonly object _gate = new();
    private readonly List<JsonObject> _envelopes = [];
    private readonly List<string> _paths = [];
    private readonly Task _serveLoop;

    internal FakeIngestionEndpoint()
    {
        Port = StartListener(_listener);
        _serveLoop = Task.Run(ServeAsync);
    }

    internal int Port { get; }

    /// <summary>Points every service of the connection string at this endpoint, so nothing can leave the machine.</summary>
    internal string ConnectionString =>
        $"InstrumentationKey={InstrumentationKey};IngestionEndpoint=http://127.0.0.1:{Port}/;LiveEndpoint=http://127.0.0.1:{Port}/";

    /// <summary>Every envelope received so far, in arrival order.</summary>
    internal IReadOnlyList<JsonObject> Envelopes
    {
        get
        {
            lock (_gate)
            {
                return [.. _envelopes];
            }
        }
    }

    /// <summary>The path of every request received so far, including requests that carried no envelopes.</summary>
    internal IReadOnlyList<string> RequestPaths
    {
        get
        {
            lock (_gate)
            {
                return [.. _paths];
            }
        }
    }

    internal IReadOnlyList<JsonObject> EnvelopesOfType(string baseType) =>
        [.. Envelopes.Where(envelope => BaseType(envelope) == baseType)];

    internal static string? BaseType(JsonObject envelope) =>
        envelope["data"]?["baseType"]?.GetValue<string>();

    /// <summary>Waits until the condition holds for the envelopes received so far, or the timeout passes.</summary>
    internal bool WaitFor(Func<IReadOnlyList<JsonObject>, bool> condition, TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (true)
        {
            if (condition(Envelopes))
            {
                return true;
            }

            if (DateTime.UtcNow >= deadline)
            {
                return false;
            }

            Thread.Sleep(20);
        }
    }

    public void Dispose()
    {
        _listener.Close();
        try
        {
            _serveLoop.Wait(TimeSpan.FromSeconds(5));
        }
        catch (AggregateException)
        {
            // The listener was closed while a request was being accepted.
        }
    }

    // Listening on loopback needs no URL reservation; retry in case another process took the port.
    private static int StartListener(HttpListener listener)
    {
        for (var attempt = 0; ; attempt++)
        {
            var port = FindFreePort();
            listener.Prefixes.Clear();
            listener.Prefixes.Add($"http://127.0.0.1:{port}/");
            try
            {
                listener.Start();
                return port;
            }
            catch (HttpListenerException) when (attempt < 10)
            {
            }
        }
    }

    private static int FindFreePort()
    {
        var probe = new TcpListener(IPAddress.Loopback, 0);
        probe.Start();
        try
        {
            return ((IPEndPoint)probe.LocalEndpoint).Port;
        }
        finally
        {
            probe.Stop();
        }
    }

    private async Task ServeAsync()
    {
        while (_listener.IsListening)
        {
            HttpListenerContext context;
            try
            {
                context = await _listener.GetContextAsync();
            }
            catch (Exception ex) when (ex is HttpListenerException or ObjectDisposedException)
            {
                return;
            }

            Respond(context);
        }
    }

    private void Respond(HttpListenerContext context)
    {
        var envelopes = new List<JsonObject>();
        try
        {
            using (var body = new MemoryStream())
            {
                context.Request.InputStream.CopyTo(body);
                envelopes.AddRange(ParseEnvelopes(Decode(body.ToArray())));
            }

            lock (_gate)
            {
                _paths.Add(context.Request.Url?.AbsolutePath ?? string.Empty);
                _envelopes.AddRange(envelopes);
            }

            var reply = Encoding.UTF8.GetBytes(
                $"{{\"itemsReceived\":{envelopes.Count},\"itemsAccepted\":{envelopes.Count},\"errors\":[]}}");
            context.Response.StatusCode = (int)HttpStatusCode.OK;
            context.Response.ContentType = "application/json";
            context.Response.OutputStream.Write(reply);
        }
        finally
        {
            context.Response.Close();
        }
    }

    // The exporter may compress its payload, so accept a gzip body whatever the headers say.
    private static string Decode(byte[] body)
    {
        if (body.Length > 2 && body[0] == 0x1f && body[1] == 0x8b)
        {
            using var gzip = new GZipStream(new MemoryStream(body), CompressionMode.Decompress);
            using var decompressed = new MemoryStream();
            gzip.CopyTo(decompressed);
            body = decompressed.ToArray();
        }

        return Encoding.UTF8.GetString(body);
    }

    // The payload is newline-delimited JSON, one envelope per line (a JSON array is accepted too).
    private static IEnumerable<JsonObject> ParseEnvelopes(string text)
    {
        if (text.TrimStart().StartsWith('[') && JsonNode.Parse(text) is JsonArray array)
        {
            foreach (var item in array.OfType<JsonObject>())
            {
                yield return item.DeepClone().AsObject();
            }

            yield break;
        }

        foreach (var line in text.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (JsonNode.Parse(line) is JsonObject envelope)
            {
                yield return envelope;
            }
        }
    }
}
