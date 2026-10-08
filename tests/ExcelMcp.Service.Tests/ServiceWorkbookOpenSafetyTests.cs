using System.Runtime.ExceptionServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "SessionLifecycle")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Medium")]
public sealed class ServiceWorkbookOpenSafetyTests
{
    [Fact]
    public async Task Open_WebPageReturnedForWorkbookPath_DoesNotPublishSession()
    {
        var path = Path.Join(Path.GetTempPath(), $"WorkbookOpenSafety_{Guid.NewGuid():N}.xlsm");
        var webPagePath = Path.ChangeExtension(path, ".mht");
        var originalHook = ExcelBatch.AfterWorkbookOpenHookForTests;
        var service = new ExcelMcpService();
        Exception? primaryFailure = null;
        var cleanupFailures = new List<Exception>();
        try
        {
            var created = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.create",
                Args = JsonSerializer.Serialize(new { filePath = path }, ServiceProtocol.JsonOptions)
            });
            Assert.True(created.Success, created.ErrorMessage);
            using var createdResult = JsonDocument.Parse(created.Result!);
            var session = createdResult.RootElement.GetProperty("sessionId").GetString();
            var closed = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.close",
                SessionId = session,
                Args = """{"save":true}"""
            });
            Assert.True(closed.Success, closed.ErrorMessage);
            ExcelBatch.AfterWorkbookOpenHookForTests = (_, book) =>
            {
                var workbook = (Excel.Workbook)book;
                workbook.SaveAs(webPagePath, Excel.XlFileFormat.xlWebArchive);
                Assert.Equal(Excel.XlFileFormat.xlWebArchive, workbook.FileFormat);
            };
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.open",
                Args = JsonSerializer.Serialize(new { filePath = path, show = false }, ServiceProtocol.JsonOptions)
            });

            Assert.False(response.Success);
            Assert.Contains("not a supported Excel workbook", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, service.SessionCount);

            ExcelBatch.AfterWorkbookOpenHookForTests = originalHook;
            var reopened = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.open",
                Args = JsonSerializer.Serialize(new { filePath = path, show = false }, ServiceProtocol.JsonOptions)
            });
            Assert.True(reopened.Success, reopened.ErrorMessage);
            using var reopenedResult = JsonDocument.Parse(reopened.Result!);
            var reopenedSession = reopenedResult.RootElement.GetProperty("sessionId").GetString();
            var inspected = await service.ProcessAsync(new ServiceRequest
            {
                Command = "workbook.get-info",
                SessionId = reopenedSession
            });
            Assert.True(inspected.Success, inspected.ErrorMessage);
            using var inspectedResult = JsonDocument.Parse(inspected.Result!);
            Assert.Equal(52, inspectedResult.RootElement.GetProperty("formatCode").GetInt32());
            Assert.Equal(path, inspectedResult.RootElement.GetProperty("fullName").GetString());
            var closedAfterRetry = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.close",
                SessionId = reopenedSession,
                Args = """{"save":false}"""
            });
            Assert.True(closedAfterRetry.Success, closedAfterRetry.ErrorMessage);
            Assert.Equal(0, service.SessionCount);
        }
        catch (Exception ex)
        {
            primaryFailure = ex;
        }
        finally
        {
            ExcelBatch.AfterWorkbookOpenHookForTests = originalHook;
            CaptureCleanup(service.Dispose);
            CaptureCleanup(() => File.Delete(path));
            CaptureCleanup(() => File.Delete(webPagePath));
        }
        if (cleanupFailures.Count > 0)
        {
            if (primaryFailure is not null) cleanupFailures.Insert(0, primaryFailure);
            throw new AggregateException("Workbook open-safety regression or cleanup failed.", cleanupFailures);
        }
        if (primaryFailure is not null) ExceptionDispatchInfo.Capture(primaryFailure).Throw();

        void CaptureCleanup(Action action)
        {
            try { action(); }
            catch (Exception ex) { cleanupFailures.Add(ex); }
        }
    }
}
