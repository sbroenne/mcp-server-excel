using System.Runtime.ExceptionServices;
using System.Text.Json;
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
    public async Task Open_HtmlFile_RetainsTheFormatExcelOpened()
    {
        var path = Path.Join(Path.GetTempPath(), $"WorkbookOpen_{Guid.NewGuid():N}.htm");
        const string contents = "<html><body><table><tr><td>native-open-marker</td></tr></table></body></html>";
        var service = new ExcelMcpService();
        Exception? primaryFailure = null;
        var cleanupFailures = new List<Exception>();
        try
        {
            await File.WriteAllTextAsync(path, contents);
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.open",
                Args = JsonSerializer.Serialize(new { filePath = path, show = false }, ServiceProtocol.JsonOptions)
            });

            Assert.True(response.Success, response.ErrorMessage);
            Assert.Equal(1, service.SessionCount);
            using var openedResult = JsonDocument.Parse(response.Result!);
            var openedSession = openedResult.RootElement.GetProperty("sessionId").GetString();
            var inspected = await service.ProcessAsync(new ServiceRequest
            {
                Command = "workbook.get-info",
                SessionId = openedSession
            });
            Assert.True(inspected.Success, inspected.ErrorMessage);
            using var inspectedResult = JsonDocument.Parse(inspected.Result!);
            Assert.Equal((int)Excel.XlFileFormat.xlHtml, inspectedResult.RootElement.GetProperty("formatCode").GetInt32());
            Assert.Equal(path, inspectedResult.RootElement.GetProperty("fullName").GetString());
            var closedAfterRetry = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.close",
                SessionId = openedSession,
                Args = """{"save":false}"""
            });
            Assert.True(closedAfterRetry.Success, closedAfterRetry.ErrorMessage);
            Assert.Equal(0, service.SessionCount);
            Assert.Equal(contents, await File.ReadAllTextAsync(path));
        }
        catch (Exception ex)
        {
            primaryFailure = ex;
        }
        finally
        {
            CaptureCleanup(service.Dispose);
            CaptureCleanup(() => File.Delete(path));
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
