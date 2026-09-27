using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacOfficeBridgeClientTests
{
    [Fact]
    public async Task Invoke_UsesExactSessionWorkbookAndReturnsTerminalValue()
    {
        var requests = new List<(string Path, JsonObject Body)>();
        using var client = CreateClient((request, body) =>
        {
            requests.Add((request.RequestUri!.AbsolutePath, body));
            return request.RequestUri.AbsolutePath switch
            {
                "/v1/sessions" => Json(HttpStatusCode.Created, """{"available":true}"""),
                "/v1/requests" => Json(HttpStatusCode.Accepted, """{"requestId":"request-1","status":"queued"}"""),
                "/v1/requests/status" => Json(HttpStatusCode.OK,
                    """{"requestId":"request-1","status":"completed","result":{"success":true,"value":{"success":true,"errorMessage":null,"action":"create"},"errorMessage":null}}"""),
                _ => throw new InvalidOperationException(request.RequestUri.AbsolutePath)
            };
        });

        var result = await client.InvokeAsync(
            "session-1",
            "/tmp/book.xlsx",
            "table.create",
            new JsonObject { ["tableName"] = "Sales" },
            TimeSpan.FromSeconds(1),
            mutation: true);

        Assert.True(result.GetProperty("success").GetBoolean());
        Assert.Equal("session-1", requests[0].Body["sessionId"]!.GetValue<string>());
        Assert.Equal("file:///tmp/book.xlsx", requests[0].Body["workbookUrl"]!.GetValue<string>());
        Assert.All(
            requests.Where(item => item.Path != "/v1/sessions"),
            item => Assert.Equal("file:///tmp/book.xlsx", item.Body["workbookUrl"]!.GetValue<string>()));
    }

    [Fact]
    public async Task Invoke_WhenAddInIsInactive_ReturnsActionableUnavailableError()
    {
        using var client = CreateClient((request, body) =>
            request.RequestUri!.AbsolutePath == "/v1/sessions"
                ? Json(HttpStatusCode.Created, """{"available":false}""")
                : Json(HttpStatusCode.Conflict,
                    """{"success":false,"errorMessage":"The Office.js add-in is not active for this workbook."}"""));

        var error = await Assert.ThrowsAsync<MacOfficeBridgeException>(() =>
            client.InvokeAsync(
                "session-1",
                "/tmp/book.xlsx",
                "table.list",
                new JsonObject(),
                TimeSpan.FromSeconds(1),
                mutation: false));

        Assert.Equal("OfficeAddInUnavailable", error.ErrorCategory);
        Assert.Contains("not active", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Invoke_WhenWorkbookBindingIsWrong_RemainsFatal()
    {
        using var client = CreateClient((request, body) =>
            request.RequestUri!.AbsolutePath == "/v1/sessions"
                ? Json(HttpStatusCode.Conflict,
                    """{"success":false,"errorMessage":"The workbook is already bound to another bridge session."}""")
                : throw new InvalidOperationException());

        var error = await Assert.ThrowsAsync<MacOfficeBridgeException>(() =>
            client.InvokeAsync(
                "session-1",
                "/tmp/book.xlsx",
                "table.list",
                new JsonObject(),
                TimeSpan.FromSeconds(1),
                mutation: false));

        Assert.Equal("WorkbookBinding", error.ErrorCategory);
    }

    [Fact]
    public async Task Invoke_WhenDispatchedMutationExpires_ReportsUncertainMutationAndCancels()
    {
        var cancelled = false;
        using var client = CreateClient((request, body) =>
        {
            return request.RequestUri!.AbsolutePath switch
            {
                "/v1/sessions" => Json(HttpStatusCode.Created, """{"available":true}"""),
                "/v1/requests" => Json(HttpStatusCode.Accepted, """{"requestId":"request-1","status":"queued"}"""),
                "/v1/requests/status" => Json(HttpStatusCode.OK,
                    """{"requestId":"request-1","status":"active","result":null}"""),
                "/v1/requests/cancel" => Cancel(),
                _ => throw new InvalidOperationException()
            };
        }, TimeSpan.FromMilliseconds(1));

        var error = await Assert.ThrowsAsync<MacOfficeMutationUncertainException>(() =>
            client.InvokeAsync(
                "session-1",
                "/tmp/book.xlsx",
                "table.create",
                new JsonObject(),
                TimeSpan.FromMilliseconds(20),
                mutation: true));

        Assert.True(cancelled);
        Assert.Contains("may still complete", error.Message, StringComparison.OrdinalIgnoreCase);

        HttpResponseMessage Cancel()
        {
            cancelled = true;
            return Json(HttpStatusCode.OK, """{"cancelled":true}""");
        }
    }

    [Fact]
    public async Task Invoke_WhenMutationExpiresBetweenPolls_UsesPersistedDispatchState()
    {
        using var client = CreateClient((request, body) =>
            request.RequestUri!.AbsolutePath switch
            {
                "/v1/sessions" => Json(HttpStatusCode.Created, """{"available":true}"""),
                "/v1/requests" => Json(HttpStatusCode.Accepted, """{"requestId":"request-1","status":"queued","dispatched":false}"""),
                "/v1/requests/status" => Json(HttpStatusCode.OK,
                    """{"requestId":"request-1","status":"expired","dispatched":true,"result":null}"""),
                _ => throw new InvalidOperationException()
            });

        var error = await Assert.ThrowsAsync<MacOfficeMutationUncertainException>(() =>
            client.InvokeAsync(
                "session-1",
                "/tmp/book.xlsx",
                "table.create",
                new JsonObject(),
                TimeSpan.FromSeconds(1),
                mutation: true));

        Assert.Contains("may still complete", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    private static MacOfficeBridgeClient CreateClient(
        Func<HttpRequestMessage, JsonObject, HttpResponseMessage> responder,
        TimeSpan? pollInterval = null)
    {
        var handler = new DelegateHandler(async request =>
        {
            var text = request.Content is null ? "{}" : await request.Content.ReadAsStringAsync();
            return responder(request, JsonNode.Parse(text)!.AsObject());
        });
        return new MacOfficeBridgeClient(
            new HttpClient(handler),
            new MacOfficeBridgeConfiguration(
                new Uri("https://localhost:47132"),
                "token",
                CertificatePath: null,
                new HashSet<string>(StringComparer.Ordinal)),
            pollInterval ?? TimeSpan.Zero);
    }

    private static HttpResponseMessage Json(HttpStatusCode statusCode, string body) =>
        new(statusCode)
        {
            Content = new StringContent(body, Encoding.UTF8, "application/json")
        };

    private sealed class DelegateHandler(
        Func<HttpRequestMessage, Task<HttpResponseMessage>> send)
        : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request,
            CancellationToken cancellationToken) => send(request);
    }
}
