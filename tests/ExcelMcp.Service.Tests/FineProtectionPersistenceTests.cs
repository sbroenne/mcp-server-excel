using System.Runtime.ExceptionServices;
using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Protection")]
[Trait("RequiresExcel", "true")]
public sealed class FineProtectionPersistenceTests
{
    [Fact]
    public async Task SaveReopen_PreservesFlagsButNotRuntimeAutomationPermission()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fine-protection-{Guid.NewGuid():N}.xlsx");
        var service = new ExcelMcpService();
        string? session = null;
        Exception? failure = null;
        try
        {
            using (var created = await Send(service, null, "session.create", new { filePath = path }))
                session = created.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrEmpty(session));
            using (await Send(service, session, "rangelink.set-cell-protection", new
            {
                sheetName = "Sheet1",
                rangeAddress = "A1",
                locked = true,
                formulaHidden = true
            })) { }
            using (await Send(service, session, "sheet.set-protection", new
            {
                sheetName = "Sheet1",
                isProtected = true,
                options = new { userInterfaceOnly = true, allowFormattingRows = true }
            })) { }
            using (var before = await Send(service, session, "sheet.get-protection", new { sheetName = "Sheet1" }))
                Assert.True(before.RootElement.GetProperty("userInterfaceOnly").GetBoolean());
            await Close(service, session!, save: true);
            session = null;
            using (var opened = await Send(service, null, "session.open", new { filePath = path }))
                session = opened.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrEmpty(session));
            using (var read = await Send(service, session, "sheet.get-protection", new { sheetName = "Sheet1" }))
            {
                Assert.True(read.RootElement.GetProperty("protectContents").GetBoolean());
                Assert.False(read.RootElement.GetProperty("userInterfaceOnly").GetBoolean());
                Assert.True(read.RootElement.GetProperty("permissions").GetProperty("allowFormattingRows").GetBoolean());
            }
            using (var cells = await Send(service, session, "rangelink.get-cell-protection", new
            {
                sheetName = "Sheet1",
                rangeAddress = "A1"
            }))
            {
                var cell = Assert.Single(cells.RootElement.GetProperty("cells").EnumerateArray());
                Assert.True(cell.GetProperty("locked").GetBoolean());
                Assert.True(cell.GetProperty("formulaHidden").GetBoolean());
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
                try
                {
                    await Close(service, session, save: false);
                }
                catch (Exception exception)
                {
                    failure = PersistentServiceCleanupFailures.Combine(failure, exception);
                }
            }
            try
            {
                service.Dispose();
            }
            catch (Exception exception)
            {
                failure = PersistentServiceCleanupFailures.Combine(failure, exception);
            }
            try
            {
                File.Delete(path);
            }
            catch (Exception exception)
            {
                failure = PersistentServiceCleanupFailures.Combine(failure, exception);
            }
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

    private static async Task Close(ExcelMcpService service, string session, bool save)
    {
        await Dispatch(service, session, "session.close", new { save });
    }

    private static async Task<ServiceResponse> Dispatch(ExcelMcpService service, string? session, string command, object args)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = command,
            SessionId = session,
            Args = JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions),
            Source = "protection-persistence-test"
        });
        Assert.True(response.Success, response.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(response.ErrorMessage));
        return response;
    }
}
