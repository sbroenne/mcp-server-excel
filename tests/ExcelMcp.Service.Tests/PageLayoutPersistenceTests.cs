using System.Runtime.ExceptionServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PageLayout")]
[Trait("RequiresExcel", "true")]
public sealed class PageLayoutPersistenceTests
{
    [Fact]
    public async Task SaveReopen_PreservesReportLayoutAndExportsExactlyTwoPdfPages()
    {
        var path = Path.Combine(Path.GetTempPath(), $"page-layout-{Guid.NewGuid():N}.xlsx");
        var pdf = Path.ChangeExtension(path, ".pdf");
        var service = new ExcelMcpService();
        string? session = null;
        Exception? failure = null;
        try
        {
            using (var created = await Send(service, null, "session.create", new { filePath = path }))
                session = created.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrEmpty(session));
            using (await Send(service, session, "range.set-values", new
            {
                sheetName = "Sheet1",
                rangeAddress = "A1:B2",
                values = new object[][] { ["Report", "Amount"], ["A", 10] }
            })) { }
            using (await Send(service, session, "sheet.set-page-setup", new
            {
                sheetName = "Sheet1",
                orientation = "landscape",
                pageSetupOptions = new
                {
                    printArea = "A1:B20",
                    printTitleRows = "1:1",
                    leftMargin = 36d,
                    centerHeader = "Saved report",
                    rightFooter = "&P",
                    paperSize = "xlPaperA4",
                    zoomPercent = 100
                }
            })) { }
            using (await Send(service, session, "sheet.set-page-breaks", new
            {
                sheetName = "Sheet1",
                pageBreakOptions = new { rows = new List<int> { 10 }, columns = new List<int>() }
            })) { }
            await Dispatch(service, session, "session.close", new { save = true });
            session = null;
            using (var opened = await Send(service, null, "session.open", new { filePath = path }))
                session = opened.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrEmpty(session));
            using (var state = await Send(service, session, "sheet.get-page-setup", new { sheetName = "Sheet1" }))
            {
                Assert.Equal("$A$1:$B$20", state.RootElement.GetProperty("printArea").GetString());
                Assert.Equal("$1:$1", state.RootElement.GetProperty("printTitleRows").GetString());
                Assert.Equal("landscape", state.RootElement.GetProperty("orientation").GetString());
                Assert.Equal(36d, state.RootElement.GetProperty("leftMargin").GetDouble(), 2);
                Assert.Equal("Saved report", state.RootElement.GetProperty("centerHeader").GetString());
                Assert.Equal("&P", state.RootElement.GetProperty("rightFooter").GetString());
                Assert.Equal("xlPaperA4", state.RootElement.GetProperty("paperSize").GetString());
                Assert.Equal(100, state.RootElement.GetProperty("zoomPercent").GetInt32());
            }
            using (var breaks = await Send(service, session, "sheet.get-page-breaks", new { sheetName = "Sheet1" }))
                Assert.Contains(breaks.RootElement.GetProperty("horizontal").EnumerateArray(),
                    item => item.GetProperty("isManual").GetBoolean() && item.GetProperty("position").GetInt32() == 10);
            using (await Send(service, session, "workbook.export-fixed-format", new { targetPath = pdf })) { }
            var bytes = File.ReadAllBytes(pdf);
            Assert.True(bytes.Length > 1000);
            var text = Encoding.Latin1.GetString(bytes);
            Assert.StartsWith("%PDF-", text, StringComparison.Ordinal);
            Assert.Equal(2, Regex.Matches(text, @"/Type\s*/Page\b", RegexOptions.CultureInvariant).Count);
            Assert.Contains("/MediaBox", text, StringComparison.Ordinal);
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
            try { File.Delete(pdf); File.Delete(path); }
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
            Source = "page-layout-test"
        });
        Assert.True(response.Success, response.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(response.ErrorMessage));
        return response;
    }
}
