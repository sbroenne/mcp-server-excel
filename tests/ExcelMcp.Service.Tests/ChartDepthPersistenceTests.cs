using System.Runtime.ExceptionServices;
using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "ChartDepth")]
[Trait("RequiresExcel", "true")]
public sealed class ChartDepthPersistenceTests
{
    [Fact]
    public async Task SaveReopen_PreservesComboAxesPointColorsAndErrorBars()
    {
        var path = Path.Combine(Path.GetTempPath(), $"chart-depth-{Guid.NewGuid():N}.xlsx");
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
                rangeAddress = "A1:C4",
                values = new object[][] { ["Category", "First", "Second"], ["A", 10, 100], ["B", 20, 200], ["C", 30, 300] }
            })) { }
            using (await Send(service, session, "chart.create-from-range", new { sheetName = "Sheet1", sourceRangeAddress = "A1:C4", chartType = "ColumnClustered", chartName = "SavedChart" })) { }
            using (await Send(service, session, "chartconfig.set-series-chart-type", new { chartName = "SavedChart", seriesIndex = 2, chartType = "LineMarkers" })) { }
            using (await Send(service, session, "chartconfig.set-series-axis-group", new { chartName = "SavedChart", seriesIndex = 2, axisGroup = "Secondary" })) { }
            using (await Send(service, session, "chartconfig.set-point-format", new { chartName = "SavedChart", seriesIndex = 1, pointIndex = 2, pointOptions = new { fillColor = "#FF0000", lineColor = "#0000FF", lineWeight = 2d } })) { }
            using (await Send(service, session, "chartconfig.set-point-format", new { chartName = "SavedChart", seriesIndex = 2, pointIndex = 1, pointOptions = new { fillColor = "#70AD47", markerStyle = "Diamond", markerSize = 14 } })) { }
            using (await Send(service, session, "chartconfig.set-error-bars", new { chartName = "SavedChart", seriesIndex = 1, errorBarOptions = new { kind = "Percent", amount = 5d, endStyle = "NoCap" } })) { }
            await Dispatch(service, session, "session.close", new { save = true });
            session = null;
            using (var opened = await Send(service, null, "session.open", new { filePath = path }))
                session = opened.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrEmpty(session));
            using (var state = await Send(service, session, "chartconfig.get-series-settings", new { chartName = "SavedChart", seriesIndex = 2 }))
            {
                Assert.Equal("LineMarkers", state.RootElement.GetProperty("chartType").GetString());
                Assert.Equal("Secondary", state.RootElement.GetProperty("axisGroup").GetString());
                Assert.Equal(3, state.RootElement.GetProperty("pointCount").GetInt32());
            }
            using (var state = await Send(service, session, "chartconfig.get-point-format", new { chartName = "SavedChart", seriesIndex = 1, pointIndex = 2 }))
            {
                Assert.Equal("#FF0000", state.RootElement.GetProperty("fillColor").GetString());
                Assert.Equal("#0000FF", state.RootElement.GetProperty("lineColor").GetString());
            }
            using (var state = await Send(service, session, "chartconfig.get-point-format", new { chartName = "SavedChart", seriesIndex = 2, pointIndex = 1 }))
            {
                Assert.Equal("#70AD47", state.RootElement.GetProperty("fillColor").GetString());
                Assert.Equal("Diamond", state.RootElement.GetProperty("markerStyle").GetString());
                Assert.Equal(14, state.RootElement.GetProperty("markerSize").GetInt32());
            }
            using (var state = await Send(service, session, "chartconfig.get-error-bars", new { chartName = "SavedChart", seriesIndex = 1 }))
            {
                Assert.True(state.RootElement.GetProperty("hasErrorBars").GetBoolean());
                Assert.Equal("NoCap", state.RootElement.GetProperty("endStyle").GetString());
            }
            using (var state = await Send(service, session, "chart.read", new { chartName = "SavedChart" }))
                Assert.Equal("Secondary", state.RootElement.GetProperty("series")[1].GetProperty("axisGroup").GetString());
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
            Source = "chart-depth-test"
        });
        Assert.True(response.Success, $"{command}: {response.ErrorMessage}");
        Assert.True(string.IsNullOrEmpty(response.ErrorMessage));
        return response;
    }
}
