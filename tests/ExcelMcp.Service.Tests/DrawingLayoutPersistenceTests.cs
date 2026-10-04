using System.Runtime.ExceptionServices;
using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "DrawingLayout")]
[Trait("RequiresExcel", "true")]
public sealed class DrawingLayoutPersistenceTests
{
    private static readonly string[] MemberTexts = ["First", "Second"];

    [Fact]
    public async Task SaveReopen_PreservesGroupedAndDuplicatedMembersGeometryAndOrder()
    {
        var path = Path.Combine(Path.GetTempPath(), $"drawing-layout-{Guid.NewGuid():N}.xlsx");
        var service = new ExcelMcpService();
        string? session = null;
        Exception? failure = null;
        try
        {
            using (var created = await Send(service, null, "session.create", new { filePath = path }))
                session = created.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrEmpty(session));
            foreach (var (name, left) in new[] { ("First", 20d), ("Second", 80d) })
                using (await Send(service, session, "drawing.add-shape", new { sheetName = "Sheet1", name, left, top = 30d, width = 40d, height = 20d, text = name })) { }
            using (await Send(service, session, "drawing.group-objects", new
            {
                sheetName = "Sheet1",
                objectNames = new List<string> { "First", "Second" },
                groupName = "Original"
            })) { }
            using (await Send(service, session, "drawing.duplicate-object", new
            {
                sheetName = "Sheet1",
                objectName = "Original",
                newName = "Copy",
                offsetLeft = 25d,
                offsetTop = 50d
            })) { }
            using (await Send(service, session, "drawing.set-z-order", new { sheetName = "Sheet1", objectName = "Copy", zOrder = "SendToBack" })) { }
            await Dispatch(service, session, "session.close", new { save = true });
            session = null;
            using (var opened = await Send(service, null, "session.open", new { filePath = path }))
                session = opened.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrEmpty(session));
            using (var state = await Send(service, session, "drawing.get-object", new { sheetName = "Sheet1", objectName = "Copy" }))
            {
                var copy = state.RootElement.GetProperty("drawingObject");
                Assert.Equal("Group", copy.GetProperty("kind").GetString());
                Assert.Equal(45d, copy.GetProperty("left").GetDouble(), 2);
                Assert.Equal(80d, copy.GetProperty("top").GetDouble(), 2);
                Assert.Equal(100d, copy.GetProperty("width").GetDouble(), 2);
                Assert.Equal(20d, copy.GetProperty("height").GetDouble(), 2);
                Assert.Equal(1, copy.GetProperty("zOrderPosition").GetInt32());
                Assert.Equal(2, copy.GetProperty("children").GetArrayLength());
                Assert.Equal(MemberTexts, copy.GetProperty("children").EnumerateArray().Select(item => item.GetProperty("text").GetString()).Order(StringComparer.Ordinal));
            }
            using (var members = await Send(service, session, "drawing.ungroup-object", new { sheetName = "Sheet1", objectName = "Copy" }))
            {
                var shapes = members.RootElement.GetProperty("drawingObjects").EnumerateArray()
                    .OrderBy(item => item.GetProperty("left").GetDouble()).ToArray();
                Assert.Equal(2, shapes.Length);
                for (var index = 0; index < shapes.Length; index++)
                {
                    Assert.Equal(MemberTexts[index], shapes[index].GetProperty("text").GetString());
                    Assert.Equal(45d + index * 60d, shapes[index].GetProperty("left").GetDouble(), 2);
                    Assert.Equal(80d, shapes[index].GetProperty("top").GetDouble(), 2);
                    Assert.Equal(40d, shapes[index].GetProperty("width").GetDouble(), 2);
                    Assert.Equal(20d, shapes[index].GetProperty("height").GetDouble(), 2);
                }
            }
            using (var objects = await Send(service, session, "drawing.list-objects", new { sheetName = "Sheet1" }))
            {
                Assert.Equal(3, objects.RootElement.GetProperty("drawingObjects").GetArrayLength());
                Assert.Contains(objects.RootElement.GetProperty("drawingObjects").EnumerateArray(), item =>
                    item.GetProperty("name").GetString() == "Original" && item.GetProperty("children").GetArrayLength() == 2);
            }
            using (var state = await Send(service, session, "drawing.get-object", new { sheetName = "Sheet1", objectName = "Original" }))
            {
                var original = state.RootElement.GetProperty("drawingObject");
                Assert.Equal(20d, original.GetProperty("left").GetDouble(), 2);
                Assert.Equal(30d, original.GetProperty("top").GetDouble(), 2);
                Assert.Equal(100d, original.GetProperty("width").GetDouble(), 2);
                Assert.Equal(20d, original.GetProperty("height").GetDouble(), 2);
                Assert.Equal(MemberTexts, original.GetProperty("children").EnumerateArray()
                    .Select(item => item.GetProperty("text").GetString()).Order(StringComparer.Ordinal));
            }
        }
        catch (Exception exception)
        {
            failure = exception;
        }
        finally
        {
            if (session is not null)
            {
                try { await Dispatch(service, session, "session.close", new { save = false }); }
                catch (Exception exception) { failure = PersistentServiceCleanupFailures.Combine(failure, exception); }
            }
            try { service.Dispose(); }
            catch (Exception exception) { failure = PersistentServiceCleanupFailures.Combine(failure, exception); }
            try { File.Delete(path); }
            catch (Exception exception) { failure = PersistentServiceCleanupFailures.Combine(failure, exception); }
        }
        if (failure is not null)
            ExceptionDispatchInfo.Capture(failure).Throw();
    }

    private static async Task<JsonDocument> Send(ExcelMcpService service, string? session, string command, object args)
    {
        var response = await Dispatch(service, session, command, args);
        Assert.NotNull(response.Result);
        var document = JsonDocument.Parse(response.Result);
        if (document.RootElement.TryGetProperty("success", out var success) && !success.GetBoolean())
        {
            document.Dispose();
            Assert.Fail($"{command} returned an unsuccessful operation.");
        }
        return document;
    }

    private static async Task<ServiceResponse> Dispatch(ExcelMcpService service, string? session, string command, object args)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = command,
            SessionId = session,
            Args = JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions),
            Source = "drawing-layout-test"
        });
        Assert.True(response.Success, $"{command}: {response.ErrorMessage}");
        Assert.True(string.IsNullOrEmpty(response.ErrorMessage));
        return response;
    }
}
