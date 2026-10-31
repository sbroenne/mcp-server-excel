using System.Runtime.ExceptionServices;
using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Window")]
[Trait("RequiresExcel", "true")]
public sealed class WindowViewPersistenceTests
{
    [Theory]
    [InlineData("A1", "A1", "$A$1", "$A$1", "$A$1")]
    [InlineData("E20", "H30:I32", "$E$20", "$H$30:$I$32", "$H$30")]
    public async Task SaveReopen_PreservesLandingSheetScrollAndSelection(
        string origin, string select, string expectedOrigin, string expectedSelection, string expectedActive)
    {
        var path = Path.Combine(Path.GetTempPath(), $"window-view-{Guid.NewGuid():N}.xlsx");
        var service = new ExcelMcpService();
        using var processes = new TestExcelProcessScope();
        string? session = null;
        Exception? failure = null;
        try
        {
            using (var created = await Send(service, null, "session.create", new { filePath = path }))
                session = created.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrWhiteSpace(session));
            using (await Send(service, session, "sheet.rename", new { oldName = "Sheet1", newName = "Dashboard" })) { }
            using (await Send(service, session, "sheet.create", new { sheetName = "Other" })) { }
            using (await Send(service, session, "window.set-view",
                new { sheetName = "Dashboard", topLeftCell = "D31", select = "H40" })) { }
            await SendAndDispose(service, session, "session.close", new { save = true });
            session = null;
            using (var opened = await Send(service, null, "session.open", new { filePath = path }))
                session = opened.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrWhiteSpace(session));
            using (var baseline = await Send(service, session, "window.get-view", new { sheetName = "Dashboard" }))
            {
                Assert.Equal("$D$31", baseline.RootElement.GetProperty("topLeftCell").GetString());
                Assert.Equal("$H$40", baseline.RootElement.GetProperty("selectedRangeAddress").GetString());
            }
            using (await Send(service, session, "window.get-view", new { sheetName = "Other" })) { }
            using (var changed = await Send(service, session, "window.set-view",
                new { sheetName = "Dashboard", topLeftCell = origin, select }))
                AssertView(changed.RootElement, expectedOrigin, expectedSelection, expectedActive);
            await SendAndDispose(service, session, "session.close", new { save = true });
            session = null;
            using (var reopened = await Send(service, null, "session.open", new { filePath = path }))
                session = reopened.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrWhiteSpace(session));
            using (var context = await Send(service, session, "window.get-context", new { }))
            {
                var window = Assert.Single(context.RootElement.GetProperty("windows").EnumerateArray());
                Assert.Equal("Dashboard", window.GetProperty("sheetName").GetString());
                Assert.Equal(expectedSelection, window.GetProperty("selectedRangeAddress").GetString());
                Assert.Equal(expectedActive, window.GetProperty("activeCellAddress").GetString());
            }
            using (var saved = await Send(service, session, "window.get-view", new { sheetName = "Dashboard" }))
                AssertView(saved.RootElement, expectedOrigin, expectedSelection, expectedActive);
        }
        catch (Exception exception)
        {
            failure = exception;
        }
        finally
        {
            if (session is not null)
            {
                try { await SendAndDispose(service, session, "session.close", new { save = false }); }
                catch (Exception exception) { failure = PersistentServiceCleanupFailures.Combine(failure, exception); }
            }
            try { service.Dispose(); }
            catch (Exception exception) { failure = PersistentServiceCleanupFailures.Combine(failure, exception); }
            if (processes.FindLeftovers() is { } leftover)
                failure = PersistentServiceCleanupFailures.Combine(failure, leftover);
            try { File.Delete(path); }
            catch (Exception exception) { failure = PersistentServiceCleanupFailures.Combine(failure, exception); }
        }
        if (failure is not null)
            ExceptionDispatchInfo.Capture(failure).Throw();
    }

    private static void AssertView(JsonElement view, string origin, string selection, string active)
    {
        Assert.Equal("Dashboard", view.GetProperty("sheetName").GetString());
        Assert.Equal(origin, view.GetProperty("topLeftCell").GetString());
        Assert.Equal("range", view.GetProperty("selectionKind").GetString());
        Assert.Equal(selection, view.GetProperty("selectedRangeAddress").GetString());
        Assert.Equal(active, view.GetProperty("activeCellAddress").GetString());
    }

    private static async Task SendAndDispose(ExcelMcpService service, string? session, string command, object args)
    {
        using (await Send(service, session, command, args)) { }
    }

    private static async Task<JsonDocument> Send(ExcelMcpService service, string? session, string command, object args)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = command,
            SessionId = session,
            Args = JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions),
            Source = "window-view-persistence-test"
        });
        Assert.True(response.Success, response.ErrorMessage);
        Assert.NotNull(response.Result);
        var document = JsonDocument.Parse(response.Result);
        if (document.RootElement.TryGetProperty("success", out var success) && !success.GetBoolean())
        {
            document.Dispose();
            Assert.Fail($"{command} returned an unsuccessful operation.");
        }
        return document;
    }
}
