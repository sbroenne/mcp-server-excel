using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    [Theory]
    [InlineData(true, true)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(false, false)]
    public async Task Close_AfterModelQuery_PersistsMacroEnabledWorkbook(bool inspectVba, bool typedColumn)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "vba-model-save.xlsm");
            var session = await CreateSessionAsync(service, path);
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, "VBA and model retained");
            if (inspectVba)
                await AssertVbaStatusAsync();

            var created = await service.ProcessAsync(new ServiceRequest
            {
                Command = "powerquery.create",
                SessionId = session,
                Args = JsonSerializer.Serialize(new
                {
                    queryName = "PublishedData",
                    mCode = typedColumn
                        ? """#table(type table [Amount = Int64.Type], {{7}, {11}})"""
                        : """#table({"Amount"}, {{7}, {11}})""",
                    loadDestination = "data-model"
                }, ServiceProtocol.JsonOptions)
            });
            RequireSuccess(created);
            using (var result = JsonDocument.Parse(created.Result!))
                Assert.Equal(2, result.RootElement.GetProperty("modelRowsLoaded").GetInt32());
            var evaluated = await service.ProcessAsync(new ServiceRequest
            {
                Command = "datamodel.evaluate",
                SessionId = session,
                Args = """{"daxQuery":"EVALUATE ROW(\"Proof\", 18)"}"""
            });
            RequireSuccess(evaluated);
            using (var result = JsonDocument.Parse(evaluated.Result!))
            {
                var root = result.RootElement;
                Assert.Equal(1, root.GetProperty("rowCount").GetInt32());
                Assert.Equal(1, root.GetProperty("columnCount").GetInt32());
                Assert.Equal("[Proof]", root.GetProperty("columns")[0].GetString());
                Assert.Equal(18, root.GetProperty("rows")[0][0].GetDouble());
            }

            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(session));
            var closed = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.close",
                SessionId = session,
                Args = """{"save":true}"""
            });
            if (!closed.Success)
            {
                Assert.Same(batch, service.SessionManager.GetSession(session));
                Assert.Equal("VBA and model retained", await ReadMarkerAsync(service, session));
            }
            RequireSuccess(closed);
            sessions.TryRemove(session, out _);
            session = await OpenSessionAsync(service, path);
            sessions[session] = 0;
            Assert.Equal("VBA and model retained", await ReadMarkerAsync(service, session));
            await AssertModelRowsAsync();

            async Task AssertVbaStatusAsync()
            {
                var response = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "vba.status",
                    SessionId = session
                });
                RequireSuccess(response);
                using var result = JsonDocument.Parse(response.Result!);
                Assert.True(result.RootElement.GetProperty("projectAccess").GetBoolean());
                Assert.Equal("Design", result.RootElement.GetProperty("mode").GetString());
                Assert.Equal("None", result.RootElement.GetProperty("protection").GetString());
            }

            async Task AssertModelRowsAsync()
            {
                var response = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "datamodel.evaluate",
                    SessionId = session,
                    Args = JsonSerializer.Serialize(new
                    {
                        daxQuery = """EVALUATE ROW("Total", SUMX('PublishedData', VALUE('PublishedData'[Amount])), "Count", COUNTROWS('PublishedData'))"""
                    }, ServiceProtocol.JsonOptions)
                });
                RequireSuccess(response);
                using var result = JsonDocument.Parse(response.Result!);
                var root = result.RootElement;
                Assert.Equal(1, root.GetProperty("rowCount").GetInt32());
                Assert.Equal(2, root.GetProperty("columnCount").GetInt32());
                Assert.Equal("[Total]", root.GetProperty("columns")[0].GetString());
                Assert.Equal("[Count]", root.GetProperty("columns")[1].GetString());
                Assert.Equal(18, root.GetProperty("rows")[0][0].GetDouble());
                Assert.Equal(2, root.GetProperty("rows")[0][1].GetDouble());
            }
        });
    }
}
