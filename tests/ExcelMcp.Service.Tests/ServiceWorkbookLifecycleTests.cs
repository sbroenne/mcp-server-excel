using System.Collections.Concurrent;
using System.Runtime.ExceptionServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "Session")]
[Trait("Speed", "Medium")]
[Trait("RequiresExcel", "true")]
public sealed partial class ServiceWorkbookLifecycleTests
{
    private static readonly bool[] CloseSaveModes = [true, false];
    [Theory]
    [InlineData("connection.list")]
    [InlineData("connection.get-properties")]
    [InlineData("connection.get-refresh-status")]
    [InlineData("connection.set-properties")]
    [InlineData("connection.refresh")]
    public async Task ExternalOlap_UnavailableBackgroundSetting_DoesNotBreakMetadata(string command)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var session = await CreateSessionAsync(service, Path.Join(directory, "olap-metadata.xlsx"));
            sessions[session] = 0;
            var batch = Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(session));
            batch.Execute((context, _) =>
            {
                Excel.Connections? connections = null;
                Excel.WorkbookConnection? connection = null;
                Excel.OLEDBConnection? oledb = null;
                try
                {
                    connections = context.Book.Connections;
                    connection = connections.Add2("OlapFixture", "Local metadata fixture",
                        "OLEDB;Provider=MSOLAP.8;Data Source=localhost;Initial Catalog=Fixture;Connect Timeout=1;",
                        "Model", Excel.XlCmdType.xlCmdCube, false, false);
                    oledb = connection.OLEDBConnection;
                    Assert.True(oledb.OLAP);
                }
                finally
                {
                    ComUtilities.Release(ref oledb);
                    ComUtilities.Release(ref connection);
                    ComUtilities.Release(ref connections);
                }
            });

            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = command,
                SessionId = session,
                Args = command switch
                {
                    "connection.list" => null,
                    "connection.set-properties" =>
                        """{"connectionName":"OlapFixture","description":"Should not change","backgroundQuery":true}""",
                    _ => """{"connectionName":"OlapFixture"}"""
                }
            });
            if (command == "connection.refresh")
            {
                // This connection has no load target; it proves capability handling, not server data refresh.
                RequireSuccess(response);
                var status = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "connection.get-refresh-status",
                    SessionId = session,
                    Args = """{"connectionName":"OlapFixture"}"""
                });
                RequireSuccess(status);
                using var state = JsonDocument.Parse(status.Result!);
                Assert.True(state.RootElement.GetProperty("supportsRefreshStatus").GetBoolean());
                Assert.False(state.RootElement.GetProperty("isRefreshing").GetBoolean());
                return;
            }
            if (command == "connection.set-properties")
            {
                Assert.False(response.Success);
                Assert.Equal("InvalidInput", response.ErrorCategory);
                Assert.Contains("OLAP", response.ErrorMessage, StringComparison.Ordinal);
                batch.Execute((context, _) =>
                {
                    Excel.Connections? connections = null;
                    Excel.WorkbookConnection? connection = null;
                    try
                    {
                        connections = context.Book.Connections;
                        connection = connections.Item("OlapFixture");
                        Assert.Equal("Local metadata fixture", connection.Description);
                    }
                    finally
                    {
                        ComUtilities.Release(ref connection);
                        ComUtilities.Release(ref connections);
                    }
                });
                var synchronous = await service.ProcessAsync(new ServiceRequest
                {
                    Command = command,
                    SessionId = session,
                    Args = """{"connectionName":"OlapFixture","backgroundQuery":false,"description":"Synchronous OLAP"}"""
                });
                RequireSuccess(synchronous);
                return;
            }
            RequireSuccess(response);
            using var json = JsonDocument.Parse(response.Result!);
            var metadata = command == "connection.list"
                ? Assert.Single(json.RootElement.GetProperty("connections").EnumerateArray())
                : json.RootElement;
            if (command == "connection.get-refresh-status")
            {
                Assert.True(metadata.GetProperty("supportsRefreshStatus").GetBoolean());
                Assert.False(metadata.GetProperty("isRefreshing").GetBoolean());
            }
            else
            {
                Assert.False(metadata.GetProperty("backgroundQuery").GetBoolean());
            }
            if (command == "connection.list")
            {
                Assert.Equal("OlapFixture", metadata.GetProperty("name").GetString());
                if (metadata.TryGetProperty("lastRefresh", out var refreshed))
                    Assert.Equal(JsonValueKind.Null, refreshed.ValueKind);
            }
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ExcelStartedRefresh_BlocksCloseAndSave_UntilNewValuesCanBePersisted(bool backgroundQuery)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "background-refresh.xlsx");
            var session = await CreateSessionAsync(service, path);
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, "Unsaved edit");
            var created = await service.ProcessAsync(new ServiceRequest
            {
                Command = "powerquery.create",
                SessionId = session,
                Args = """{"queryName":"DelayedLoad","mCode":"#table({\"Marker\"}, {{\"Before\"}})","loadDestination":"worksheet","targetSheet":"DelayedLoad"}"""
            });
            RequireSuccess(created);
            var updated = await service.ProcessAsync(new ServiceRequest
            {
                Command = "powerquery.update",
                SessionId = session,
                Args = """{"queryName":"DelayedLoad","mCode":"Function.InvokeAfter(() => #table({\"Marker\"}, {{\"After\"}}), #duration(0,0,0,12))","refresh":false}"""
            });
            RequireSuccess(updated);
            var batch = Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(session));
            using var refreshEntered = new ManualResetEventSlim();
            string? refreshConnectionName = null;
            var refreshTask = Task.Run(() => batch.Execute((context, _) =>
            {
                Excel.Sheets? sheets = null;
                Excel.Worksheet? sheet = null;
                Excel.ListObjects? tables = null;
                Excel.ListObject? table = null;
                Excel.QueryTable? query = null;
                Excel.WorkbookConnection? connection = null;
                try
                {
                    sheets = context.Book.Worksheets;
                    sheet = (Excel.Worksheet)sheets["DelayedLoad"];
                    tables = sheet.ListObjects;
                    table = tables[1];
                    query = table.QueryTable;
                    connection = query.WorkbookConnection;
                    refreshConnectionName = connection.Name;
                    refreshEntered.Set();
                    Assert.True(query.Refresh(backgroundQuery));
                    if (backgroundQuery)
                    {
                        Assert.True(query.Refreshing, "The native background query must still be running.");
                    }
                }
                finally
                {
                    ComUtilities.Release(ref connection);
                    ComUtilities.Release(ref query);
                    ComUtilities.Release(ref table);
                    ComUtilities.Release(ref tables);
                    ComUtilities.Release(ref sheet);
                    ComUtilities.Release(ref sheets);
                }
            }));
            Assert.True(refreshEntered.Wait(TimeSpan.FromSeconds(10)), "Native refresh did not start.");
            if (backgroundQuery)
            {
                await refreshTask.WaitAsync(TimeSpan.FromSeconds(10));
                Assert.NotEqual(WorkbookRefreshState.Ready,
                    Assert.IsAssignableFrom<IExcelBatchRefreshState>(batch).GetRefreshState());
            }
            else
            {
                Assert.False(refreshTask.IsCompleted, "The delayed native refresh already completed.");
            }
            try
            {
                Assert.Equal(0, service.SessionManager.GetActiveOperationCount(session));
                var guardTimer = System.Diagnostics.Stopwatch.StartNew();
                var list = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
                RequireSuccess(list);
                using var json = JsonDocument.Parse(list.Result!);
                var listed = Assert.Single(json.RootElement.GetProperty("sessions").EnumerateArray());
                Assert.False(listed.GetProperty("canClose").GetBoolean());
                foreach (var save in CloseSaveModes)
                {
                    var rejected = await service.ProcessAsync(new ServiceRequest
                    {
                        Command = "session.close",
                        SessionId = session,
                        Args = JsonSerializer.Serialize(new { save }, ServiceProtocol.JsonOptions)
                    });
                    Assert.False(rejected.Success);
                    Assert.Contains("refresh", rejected.ErrorMessage, StringComparison.OrdinalIgnoreCase);
                    Assert.Same(batch, service.SessionManager.GetSession(session));
                }
                var saveFailure = Assert.Throws<ExcelBusyException>(() => batch.Save());
                Assert.Contains("refresh", saveFailure.Message, StringComparison.OrdinalIgnoreCase);
                Assert.False(batch.HasTimedOutOperation);
                Assert.True(guardTimer.Elapsed < TimeSpan.FromSeconds(5),
                    "Status, rejected close, and rejected save must not wait behind the long refresh.");
                var saveAsPath = Path.Join(directory, "must-not-save-during-refresh.xlsx");
                var saveAs = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "workbook.save-as",
                    SessionId = session,
                    Args = JsonSerializer.Serialize(new { targetPath = saveAsPath }, ServiceProtocol.JsonOptions)
                });
                Assert.False(saveAs.Success);
                Assert.Equal("Busy", saveAs.ErrorCategory);
                Assert.False(File.Exists(saveAsPath));
                Assert.Equal(path, batch.WorkbookPath);
                var status = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "connection.get-refresh-status",
                    SessionId = session,
                    Args = JsonSerializer.Serialize(
                        new { connectionName = refreshConnectionName }, ServiceProtocol.JsonOptions)
                });
                if (backgroundQuery)
                {
                    RequireSuccess(status);
                    using var state = JsonDocument.Parse(status.Result!);
                    Assert.True(state.RootElement.GetProperty("isRefreshing").GetBoolean());
                }
                else
                {
                    Assert.False(status.Success);
                    Assert.Equal("Busy", status.ErrorCategory);
                    var cancel = await service.ProcessAsync(new ServiceRequest
                    {
                        Command = "connection.cancel-refresh",
                        SessionId = session,
                        Args = JsonSerializer.Serialize(
                            new { connectionName = refreshConnectionName }, ServiceProtocol.JsonOptions)
                    });
                    Assert.False(cancel.Success);
                    Assert.Equal("Busy", cancel.ErrorCategory);
                }
                Assert.False(batch.HasTimedOutOperation);
                Assert.True(guardTimer.Elapsed < TimeSpan.FromSeconds(5),
                    "Save As and refresh controls must also reject busy access promptly.");
            }
            finally
            {
                await refreshTask.WaitAsync(TimeSpan.FromMinutes(1));
                if (backgroundQuery)
                {
                    var refreshState = Assert.IsAssignableFrom<IExcelBatchRefreshState>(batch);
                    var deadline = DateTime.UtcNow.AddMinutes(1);
                    while (refreshState.GetRefreshState() != WorkbookRefreshState.Ready && DateTime.UtcNow < deadline)
                    {
                        await Task.Delay(250);
                    }
                    Assert.Equal(WorkbookRefreshState.Ready, refreshState.GetRefreshState());
                }
            }

            Assert.False(batch.Execute((context, _) => context.Book.Saved));
            Assert.Equal("After", await ReadMarkerAsync(service, session, "DelayedLoad", "A2"));
            Assert.Equal("Unsaved edit", await ReadMarkerAsync(service, session));
            await CloseSessionAsync(service, session, save: true);
            sessions.TryRemove(session, out _);
            var reopened = await OpenSessionAsync(service, path);
            sessions[reopened] = 0;
            Assert.Equal("After", await ReadMarkerAsync(service, reopened, "DelayedLoad", "A2"));
            Assert.Equal("Unsaved edit", await ReadMarkerAsync(service, reopened));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RefreshStartsAfterClosePreflight_RetainsUsableSessionAndPersistsRetry(bool save)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "late-close-refresh.xlsx");
            var session = await CreateSessionAsync(service, path);
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, "Before refused close");
            RequireSuccess(await service.ProcessAsync(new ServiceRequest
            {
                Command = "powerquery.create",
                SessionId = session,
                Args = """{"queryName":"LateRefresh","mCode":"#table({\"Marker\"}, {{\"Before\"}})","loadDestination":"worksheet","targetSheet":"LateRefresh"}"""
            }));
            RequireSuccess(await service.ProcessAsync(new ServiceRequest
            {
                Command = "powerquery.update",
                SessionId = session,
                Args = """{"queryName":"LateRefresh","mCode":"Function.InvokeAfter(() => #table({\"Marker\"}, {{\"After\"}}), #duration(0,0,0,12))","refresh":false}"""
            }));
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(session));
            var refreshStarted = false;
            service.SessionManager.BeforeCloseSessionHookForTests = () =>
                batch.Execute((context, _) =>
                {
                    Excel.Sheets? sheets = null;
                    Excel.Worksheet? sheet = null;
                    Excel.ListObjects? tables = null;
                    Excel.ListObject? table = null;
                    Excel.QueryTable? query = null;
                    try
                    {
                        sheets = context.Book.Worksheets;
                        sheet = (Excel.Worksheet)sheets["LateRefresh"];
                        tables = sheet.ListObjects;
                        table = tables[1];
                        query = table.QueryTable;
                        Assert.True(query.Refresh(true));
                        Assert.True(query.Refreshing, "Native refresh must start after close preflight.");
                        refreshStarted = true;
                    }
                    finally
                    {
                        ComUtilities.Release(ref query);
                        ComUtilities.Release(ref table);
                        ComUtilities.Release(ref tables);
                        ComUtilities.Release(ref sheet);
                        ComUtilities.Release(ref sheets);
                    }
                });
            try
            {
                var response = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.close",
                    SessionId = session,
                    Args = JsonSerializer.Serialize(new { save }, ServiceProtocol.JsonOptions)
                });
                Assert.True(refreshStarted);
                Assert.False(response.Success);
                Assert.Equal("Busy", response.ErrorCategory);
                Assert.Same(batch, service.SessionManager.GetSession(session));
                Assert.False(batch.HasTimedOutOperation);
                Assert.True(batch.IsExcelProcessAlive());
                Assert.True(service.SessionManager.TryGetFilePath(session, out var retainedPath));
                Assert.Equal(path, retainedPath);
                Assert.Equal("Before refused close", await ReadMarkerAsync(service, session));
                await WriteMarkerAsync(service, session, "Usable after refused close");
            }
            finally
            {
                service.SessionManager.BeforeCloseSessionHookForTests = null;
                if (service.SessionManager.GetSession(session) != null)
                {
                    var deadline = DateTime.UtcNow.AddMinutes(1);
                    while (batch.GetRefreshState() != WorkbookRefreshState.Ready && DateTime.UtcNow < deadline)
                        await Task.Delay(250);
                    Assert.Equal(WorkbookRefreshState.Ready, batch.GetRefreshState());
                }
            }
            Assert.Equal("After", await ReadMarkerAsync(service, session, "LateRefresh", "A2"));
            await CloseSessionAsync(service, session, save: true);
            sessions.TryRemove(session, out _);
            var reopened = await OpenSessionAsync(service, path);
            sessions[reopened] = 0;
            Assert.Equal("Usable after refused close", await ReadMarkerAsync(service, reopened));
            Assert.Equal("After", await ReadMarkerAsync(service, reopened, "LateRefresh", "A2"));
        });
    }

    [Fact]
    public async Task GetInfo_ReturnsNativeAutoSaveStatus()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var session = await CreateSessionAsync(service, Path.Join(directory, "autosave.xlsx"));
            sessions[session] = 0;
            var batch = service.SessionManager.GetSession(session);
            Assert.NotNull(batch);
            var nativeAutoSave = batch.Execute((context, _) => context.Book.AutoSaveOn);

            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "workbook.get-info",
                SessionId = session
            });
            RequireSuccess(response);
            using var json = JsonDocument.Parse(response.Result!);
            Assert.True(json.RootElement.TryGetProperty("autoSaveOn", out var autoSave));
            Assert.Equal(nativeAutoSave, autoSave.GetBoolean());
        });
    }

    [Theory]
    [InlineData("create", false)]
    [InlineData("create", true)]
    [InlineData("open", false)]
    [InlineData("open", true)]
    public async Task CreateAndOpen_ListReportsRequestedVisibilityAndCloseState(
        string action,
        bool show)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var workbookPath = Path.Join(directory, $"{action}-{show}.xlsx");
            if (action == "open")
            {
                File.Copy(
                    Path.Join(AppContext.BaseDirectory, "TestFiles", "batch-test-static.xlsx"),
                    workbookPath);
            }

            var sessionId = action == "create"
                ? await CreateSessionAsync(service, workbookPath, show)
                : await OpenSessionAsync(service, workbookPath, show);
            sessions[sessionId] = 0;

            var list = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.list"
            });

            RequireSuccess(list);
            using var result = JsonDocument.Parse(list.Result!);
            var session = Assert.Single(
                result.RootElement.GetProperty("sessions").EnumerateArray(),
                item => item.GetProperty("sessionId").GetString() == sessionId);
            Assert.Equal(show, session.GetProperty("isExcelVisible").GetBoolean());
            Assert.Equal(0, session.GetProperty("activeOperations").GetInt32());
            Assert.True(session.GetProperty("canClose").GetBoolean());
            await CloseSessionAsync(service, sessionId, save: false);
            sessions.TryRemove(sessionId, out _);
            await AssertSessionIdsAsync(service);
        });
    }

    [Fact]
    public async Task SaveCloseReopen_PersistsMarkerInNewSession()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var workbookPath = Path.Join(directory, "persist.xlsx");
            var sessionId = await CreateSessionAsync(service, workbookPath);
            sessions[sessionId] = 0;
            await WriteMarkerAsync(service, sessionId, "Persisted");
            Assert.Equal("Persisted", await ReadMarkerAsync(service, sessionId));
            await CloseSessionAsync(service, sessionId, save: true);
            sessions.TryRemove(sessionId, out _);
            await AssertSessionIdsAsync(service);

            var reopenedSessionId = await OpenSessionAsync(service, workbookPath);
            sessions[reopenedSessionId] = 0;
            Assert.NotEqual(sessionId, reopenedSessionId);
            Assert.Equal(
                "Persisted",
                await ReadMarkerAsync(service, reopenedSessionId));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SlowReadinessInspection_AllowsSaveAndCloseWithPersistedEdits(bool saveAs)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "slow-readiness.xlsx");
            var targetPath = Path.Join(directory, "slow-readiness-saved-as.xlsx");
            var session = await CreateSessionAsync(service, path);
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, "Slow inspection persisted");
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(session));
            var inspections = 0;
            batch.BeforeRefreshStateReadHookForTests = () =>
            {
                Interlocked.Increment(ref inspections);
                Thread.Sleep(TimeSpan.FromSeconds(1.5));
            };
            try
            {
                Assert.Equal(WorkbookRefreshState.Ready, batch.GetRefreshState());
                Assert.True(service.SessionManager.ValidateClose(session).CanClose);
                if (saveAs)
                {
                    var response = await service.ProcessAsync(new ServiceRequest
                    {
                        Command = "workbook.save-as",
                        SessionId = session,
                        Args = JsonSerializer.Serialize(new { targetPath }, ServiceProtocol.JsonOptions)
                    });
                    RequireSuccess(response);
                    Assert.Equal(targetPath, batch.WorkbookPath, ignoreCase: true);
                    Assert.True(batch.Execute((context, _) => context.Book.Saved));
                }
                await CloseSessionAsync(service, session, save: !saveAs);
                sessions.TryRemove(session, out _);
                Assert.True(inspections >= 3, "Save and close must inspect actual Excel readiness.");
                Assert.False(batch.HasTimedOutOperation);
            }
            finally
            {
                batch.BeforeRefreshStateReadHookForTests = null;
                if (service.SessionManager.GetSession(session) is not null)
                    batch.Execute((_, _) => 0);
            }

            var reopened = await OpenSessionAsync(service, saveAs ? targetPath : path);
            sessions[reopened] = 0;
            Assert.Equal("Slow inspection persisted", await ReadMarkerAsync(service, reopened));
        });
    }

    [Theory]
    [InlineData("range.set-values")]
    [InlineData("sheet.create")]
    public async Task ReadOnlyWorkbook_RejectsWritesAndRetainsSavedContents(string command)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "read-only-writes.xlsx");
            var sessionId = await CreateSessionAsync(service, path);
            sessions[sessionId] = 0;
            await WriteMarkerAsync(service, sessionId, "Saved baseline");
            await CloseSessionAsync(service, sessionId, save: true);
            sessions.TryRemove(sessionId, out _);
            sessionId = await OpenSessionAsync(service, path);
            sessions[sessionId] = 0;
            var batch = Assert.IsAssignableFrom<IExcelBatch>(
                service.SessionManager.GetSession(sessionId));
            batch.Execute((context, _) =>
                context.Book.ChangeFileAccess(Excel.XlFileAccess.xlReadOnly, Type.Missing, false));
            Assert.True(batch.Execute((context, _) => context.Book.ReadOnly));
            Assert.True(batch.Execute((context, _) => context.Book.Saved));
            var beforeSheets = await service.ProcessAsync(new ServiceRequest
            {
                Command = "sheet.list",
                SessionId = sessionId
            });
            RequireSuccess(beforeSheets);
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = command,
                SessionId = sessionId,
                Args = command == "sheet.create"
                    ? """{"sheetName":"RejectedSheet"}"""
                    : """{"sheetName":"Sheet1","rangeAddress":"A1","overwritePolicy":"allow","values":[["Rejected marker"]]}"""
            });

            Assert.False(response.Success);
            Assert.Contains("read-only", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Same(batch, service.SessionManager.GetSession(sessionId));
            Assert.True(batch.Execute((context, _) => context.Book.Saved));
            Assert.Equal("Saved baseline", await ReadMarkerAsync(service, sessionId));
            var afterSheets = await service.ProcessAsync(new ServiceRequest
            {
                Command = "sheet.list",
                SessionId = sessionId
            });
            RequireSuccess(afterSheets);
            Assert.Equal(beforeSheets.Result, afterSheets.Result);

            var save = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.close",
                SessionId = sessionId,
                Args = """{"save":true}"""
            });
            Assert.False(save.Success);
            Assert.Contains("read-only", save.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Same(batch, service.SessionManager.GetSession(sessionId));
            await CloseSessionAsync(service, sessionId, save: false);
            sessions.TryRemove(sessionId, out _);
            var reopened = await OpenSessionAsync(service, path);
            sessions[reopened] = 0;
            Assert.Equal("Saved baseline", await ReadMarkerAsync(service, reopened));
        });
    }

    [Theory]
    [InlineData("session.close")]
    [InlineData("workbook.save-as")]
    public async Task CancelledSave_ReturnsFailureAndRetainsUnsavedSession(string command)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "cancelled-save.xlsx");
            var targetPath = Path.Join(directory, "cancelled-save-as.xlsx");
            var sessionId = await CreateSessionAsync(service, path);
            sessions[sessionId] = 0;
            await WriteMarkerAsync(service, sessionId, "Saved baseline");
            await CloseSessionAsync(service, sessionId, save: true);
            sessions.TryRemove(sessionId, out _);
            sessionId = await OpenSessionAsync(service, path);
            sessions[sessionId] = 0;
            await WriteMarkerAsync(service, sessionId, "Unsaved marker");
            var batch = Assert.IsAssignableFrom<IExcelBatch>(
                service.SessionManager.GetSession(sessionId));
            Assert.False(batch.Execute((context, _) => context.Book.ReadOnly));
            Assert.False(batch.Execute((context, _) => context.Book.Saved));
            var eventCount = 0;
            Excel.AppEvents_WorkbookBeforeSaveEventHandler cancelSave =
                (Excel.Workbook _, bool _, ref bool cancel) =>
                {
                    eventCount++;
                    cancel = true;
                };
            batch.Execute((context, _) => context.App.WorkbookBeforeSave += cancelSave);
            try
            {
                var response = await service.ProcessAsync(new ServiceRequest
                {
                    Command = command,
                    SessionId = sessionId,
                    Args = command == "session.close"
                        ? """{"save":true}"""
                        : JsonSerializer.Serialize(new { targetPath }, ServiceProtocol.JsonOptions)
                });

                Assert.True(eventCount > 0, "Excel did not reach its before-save event.");
                Assert.False(response.Success);
                Assert.Contains("not saved", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("remain", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
                Assert.Same(batch, service.SessionManager.GetSession(sessionId));
                Assert.False(batch.Execute((context, _) => context.Book.Saved));
                Assert.Equal(path, batch.Execute((context, _) => context.Book.FullName), ignoreCase: true);
                Assert.Equal(path, batch.WorkbookPath, ignoreCase: true);
                Assert.Equal("Unsaved marker", await ReadMarkerAsync(service, sessionId));
                Assert.False(File.Exists(targetPath));
            }
            finally
            {
                if (service.SessionManager.GetSession(sessionId) is not null)
                {
                    batch.Execute((context, _) => context.App.WorkbookBeforeSave -= cancelSave);
                }
            }

            await CloseSessionAsync(service, sessionId, save: false);
            sessions.TryRemove(sessionId, out _);
            var reopened = await OpenSessionAsync(service, path);
            sessions[reopened] = 0;
            Assert.Equal("Saved baseline", await ReadMarkerAsync(service, reopened));
        });
    }

    [Fact]
    public async Task SaveAs_AfterSaveEdits_ReturnsFailureAndTracksNewPath()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "before-save-as.xlsx");
            var targetPath = Path.Join(directory, "after-save-as.xlsx");
            var sessionId = await CreateSessionAsync(service, path);
            sessions[sessionId] = 0;
            await WriteMarkerAsync(service, sessionId, "Saved baseline");
            await CloseSessionAsync(service, sessionId, save: true);
            sessions.TryRemove(sessionId, out _);
            sessionId = await OpenSessionAsync(service, path);
            sessions[sessionId] = 0;
            var batch = Assert.IsAssignableFrom<IExcelBatch>(
                service.SessionManager.GetSession(sessionId));
            var eventCount = 0;
            Excel.AppEvents_WorkbookAfterSaveEventHandler dirtyAfterSave =
                (Excel.Workbook workbook, bool success) =>
                {
                    Assert.True(success);
                    eventCount++;
                    workbook.Saved = false;
                };
            batch.Execute((context, _) => context.App.WorkbookAfterSave += dirtyAfterSave);
            try
            {
                var response = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "workbook.save-as",
                    SessionId = sessionId,
                    Args = JsonSerializer.Serialize(new { targetPath }, ServiceProtocol.JsonOptions)
                });

                Assert.Equal(1, eventCount);
                Assert.False(response.Success);
                Assert.Contains("not saved", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
                Assert.Same(batch, service.SessionManager.GetSession(sessionId));
                Assert.False(batch.Execute((context, _) => context.Book.Saved));
                Assert.Equal(targetPath, batch.Execute((context, _) => context.Book.FullName), ignoreCase: true);
                Assert.Equal(targetPath, batch.WorkbookPath, ignoreCase: true);
                Assert.True(service.SessionManager.TryGetFilePath(sessionId, out var trackedPath));
                Assert.Equal(targetPath, trackedPath, ignoreCase: true);
                var duplicate = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.open",
                    Args = JsonSerializer.Serialize(new { filePath = targetPath }, ServiceProtocol.JsonOptions)
                });
                Assert.False(duplicate.Success);
                Assert.Contains("already open", duplicate.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            }
            finally
            {
                batch.Execute((context, _) => context.App.WorkbookAfterSave -= dirtyAfterSave);
            }

            await CloseSessionAsync(service, sessionId, save: false);
            sessions.TryRemove(sessionId, out _);
            foreach (var savedPath in new[] { path, targetPath })
            {
                var reopened = await OpenSessionAsync(service, savedPath);
                sessions[reopened] = 0;
                Assert.Equal("Saved baseline", await ReadMarkerAsync(service, reopened));
                await CloseSessionAsync(service, reopened, save: false);
                sessions.TryRemove(reopened, out _);
            }
        });
    }

    [Fact]
    public async Task CloseWithoutSaving_DiscardsEditsAndPreservesOtherWorkbook()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "discard.xlsx");
            var sessionId = await CreateSessionAsync(service, path);
            sessions[sessionId] = 0;
            await WriteMarkerAsync(service, sessionId, "Saved");
            Assert.Equal("Saved", await ReadMarkerAsync(service, sessionId));
            await CloseSessionAsync(service, sessionId, save: true);
            sessions.TryRemove(sessionId, out _);
            sessionId = await OpenSessionAsync(service, path);
            sessions[sessionId] = 0;
            Assert.Equal("Saved", await ReadMarkerAsync(service, sessionId));
            var neighbor = await CreateSessionAsync(service, Path.Join(directory, "neighbor.xlsx"));
            sessions[neighbor] = 0;
            await WriteMarkerAsync(service, neighbor, "Retained neighbor");

            await WriteMarkerAsync(service, sessionId, "Discarded");
            Assert.Equal("Discarded", await ReadMarkerAsync(service, sessionId));
            await CloseSessionAsync(service, sessionId, save: false);
            sessions.TryRemove(sessionId, out _);
            await AssertSessionIdsAsync(service, neighbor);
            Assert.Equal("Retained neighbor", await ReadMarkerAsync(service, neighbor));

            var reopened = await OpenSessionAsync(service, path);
            sessions[reopened] = 0;
            Assert.NotEqual(sessionId, reopened);
            Assert.Equal("Saved", await ReadMarkerAsync(service, reopened));
            Assert.Equal("Retained neighbor", await ReadMarkerAsync(service, neighbor));
        });
    }

    [Fact]
    public async Task CloseMissingSession_PreservesLiveWorkbookAndAllowsRecovery()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var sessionId = await CreateSessionAsync(service, Path.Join(directory, "retained.xlsx"));
            sessions[sessionId] = 0;
            await WriteMarkerAsync(service, sessionId, "Retained");
            Assert.Equal("Retained", await ReadMarkerAsync(service, sessionId));
            var batch = service.SessionManager.GetSession(sessionId);
            var missing = $"missing-{Guid.NewGuid():N}";
            var rejected = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.close",
                SessionId = missing,
                Args = """{"save":true}"""
            });
            Assert.False(rejected.Success);
            Assert.Equal("SessionNotFound", rejected.ErrorCategory);
            Assert.Equal(missing, rejected.SessionId);
            Assert.Contains("not found", rejected.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            await AssertSessionIdsAsync(service, sessionId);
            Assert.Same(batch, service.SessionManager.GetSession(sessionId));
            Assert.Equal("Retained", await ReadMarkerAsync(service, sessionId));
            await WriteMarkerAsync(service, sessionId, "Recovered");
            Assert.Equal("Recovered", await ReadMarkerAsync(service, sessionId));
            await CloseSessionAsync(service, sessionId, save: true);
            sessions.TryRemove(sessionId, out _);
            await AssertSessionIdsAsync(service);
        });
    }

    [Fact]
    public async Task ConcurrentWorkbookWorkflows_StayIsolatedAndPersist()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            const int workflowCount = 4;
            var openedCount = 0;
            var allOpened = new TaskCompletionSource<bool>(
                TaskCreationOptions.RunContinuationsAsynchronously);
            var releaseWorkflows = new TaskCompletionSource<bool>(
                TaskCreationOptions.RunContinuationsAsynchronously);
            var startupFailure = new TaskCompletionSource<Exception>(
                TaskCreationOptions.RunContinuationsAsynchronously);
            var workflows = Enumerable.Range(0, workflowCount).Select(index => Task.Run(async () =>
            {
                var workbookPath = Path.Join(directory, $"concurrent-{index}.xlsx");
                var sheetName = $"Data{index}";
                var marker = $"Marker-{index}";
                try
                {
                    var sessionId = await CreateSessionAsync(service, workbookPath);
                    sessions[sessionId] = 0;
                    if (Interlocked.Increment(ref openedCount) == workflowCount)
                    {
                        allOpened.TrySetResult(true);
                    }

                    await releaseWorkflows.Task;
                    await CreateSheetAsync(service, sessionId, sheetName);
                    await WriteWorkflowValuesAsync(service, sessionId, sheetName, marker, index);
                    Assert.Equal(marker, await ReadMarkerAsync(service, sessionId, sheetName));
                    await FormatWorkflowValuesAsync(service, sessionId, sheetName);
                    var saveState = CaptureSaveState(service, sessionId, workbookPath);
                    Assert.True(saveState.FileExists);
                    Assert.True(saveState.FileLength > 0);
                    Assert.False(saveState.ReadOnly);
                    Assert.False(saveState.HasReadOnlyAttribute);
                    Assert.False(saveState.Saved);
                    Assert.True(saveState.ProcessAlive);
                    Assert.NotNull(saveState.ExcelProcessId);
                    Assert.Equal(
                        Path.GetFullPath(workbookPath),
                        saveState.ExcelFullName,
                        ignoreCase: true);
                    await CloseSessionAsync(service, sessionId, save: true, saveState: saveState);
                    sessions.TryRemove(sessionId, out _);

                    Assert.True(File.Exists(workbookPath), $"Expected workbook to exist: {workbookPath}");
                    var reopenedSessionId = await OpenSessionAsync(service, workbookPath);
                    sessions[reopenedSessionId] = 0;
                    Assert.NotEqual(sessionId, reopenedSessionId);
                    var persisted = await ReadMarkerAsync(service, reopenedSessionId, sheetName);
                    await CloseSessionAsync(service, reopenedSessionId, save: false);
                    sessions.TryRemove(reopenedSessionId, out _);
                    return new WorkflowResult(
                        index,
                        workbookPath,
                        persisted,
                        saveState.ExcelProcessId!.Value);
                }
                catch (Exception ex)
                {
                    startupFailure.TrySetResult(ex);
                    throw;
                }
            })).ToArray();

            Exception? workflowFailure = null;
            WorkflowResult[]? results = null;
            try
            {
                var readiness = await Task.WhenAny(allOpened.Task, startupFailure.Task)
                    .WaitAsync(TimeSpan.FromMinutes(2));
                if (readiness == startupFailure.Task)
                {
                    throw await startupFailure.Task;
                }

                Assert.Equal(workflowCount, service.SessionCount);
                var list = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
                RequireSuccess(list);
                using (var result = JsonDocument.Parse(list.Result!))
                {
                    var openSessionIds = result.RootElement
                        .GetProperty("sessions")
                        .EnumerateArray()
                        .Select(item => item.GetProperty("sessionId").GetString())
                        .ToArray();
                    Assert.Equal(workflowCount, openSessionIds.Length);
                    Assert.All(sessions.Keys, sessionId => Assert.Contains(sessionId, openSessionIds));
                }
            }
            catch (Exception ex)
            {
                workflowFailure = ex;
            }
            finally
            {
                releaseWorkflows.TrySetResult(true);
                try
                {
                    results = await Task.WhenAll(workflows);
                }
                catch (Exception ex)
                {
                    workflowFailure = PersistentServiceCleanupFailures.Combine(
                        workflowFailure,
                        ex);
                }
            }

            if (workflowFailure is not null)
            {
                throw workflowFailure;
            }

            Assert.NotNull(results);
            Assert.Equal(workflowCount, results.Length);
            Assert.Equal(
                workflowCount,
                results.Select(result => result.FilePath).Distinct(StringComparer.OrdinalIgnoreCase).Count());
            Assert.Equal(
                workflowCount,
                results.Select(result => result.ExcelProcessId).Distinct().Count());
            Assert.All(
                results,
                result => Assert.Equal($"Marker-{result.Index}", result.PersistedValue));

            Assert.Equal(0, service.SessionCount);
            var finalList = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
            RequireSuccess(finalList);
            using var finalResult = JsonDocument.Parse(finalList.Result!);
            Assert.Empty(finalResult.RootElement.GetProperty("sessions").EnumerateArray());
        });
    }

    private static async Task RunWithCleanupAsync(
        Func<ExcelMcpService, string, ConcurrentDictionary<string, byte>, Task> test)
    {
        var directory = Path.Join(
            Path.GetTempPath(),
            $"ServiceWorkbookLifecycleTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        var sessions = new ConcurrentDictionary<string, byte>();
        var service = new ExcelMcpService();
        Exception? failure = null;

        try
        {
            await test(service, directory, sessions);
        }
        catch (Exception ex)
        {
            failure = ex;
        }
        finally
        {
            foreach (var sessionId in sessions.Keys)
            {
                try
                {
                    await CloseSessionAsync(service, sessionId, save: false);
                }
                catch (Exception ex)
                {
                    failure = PersistentServiceCleanupFailures.Combine(failure, ex);
                }
            }

            try
            {
                service.Dispose();
            }
            catch (Exception ex)
            {
                failure = PersistentServiceCleanupFailures.Combine(failure, ex);
            }

            try
            {
                Directory.Delete(directory, recursive: true);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                failure = PersistentServiceCleanupFailures.Combine(failure, ex);
            }
        }

        if (failure is not null)
        {
            ExceptionDispatchInfo.Capture(failure).Throw();
        }
    }

    private static async Task<string> CreateSessionAsync(
        ExcelMcpService service,
        string workbookPath,
        bool show = false)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.create",
            Args = JsonSerializer.Serialize(new
            {
                filePath = workbookPath,
                show,
                timeoutSeconds = 120
            }, ServiceProtocol.JsonOptions)
        });
        return GetSessionId(response);
    }

    private static async Task<string> OpenSessionAsync(
        ExcelMcpService service,
        string workbookPath,
        bool show = false)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new
            {
                filePath = workbookPath,
                show,
                timeoutSeconds = 120
            }, ServiceProtocol.JsonOptions)
        });
        return GetSessionId(response);
    }

    private static async Task WriteMarkerAsync(
        ExcelMcpService service,
        string sessionId,
        string marker)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                sheetName = "Sheet1",
                rangeAddress = "A1",
                overwritePolicy = "allow",
                values = new object?[][] { [marker] }
            }, ServiceProtocol.JsonOptions)
        });
        RequireSuccess(response);
    }

    private static async Task CreateSheetAsync(
        ExcelMcpService service,
        string sessionId,
        string sheetName)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "sheet.create",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new { sheetName }, ServiceProtocol.JsonOptions)
        });
        RequireSuccess(response);
    }

    private static async Task WriteWorkflowValuesAsync(
        ExcelMcpService service,
        string sessionId,
        string sheetName,
        string marker,
        int index)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                sheetName,
                rangeAddress = "A1:A2",
                overwritePolicy = "allow",
                values = new object?[][] { [marker], [$"File-{index}"] }
            }, ServiceProtocol.JsonOptions)
        });
        RequireSuccess(response);
    }

    private static async Task FormatWorkflowValuesAsync(
        ExcelMcpService service,
        string sessionId,
        string sheetName)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "rangeformat.format",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                sheetName,
                rangeAddresses = (string[])["A1:A2"],
                formatOptions = new
                {
                    bold = true
                }
            }, ServiceProtocol.JsonOptions)
        });
        RequireSuccess(response);
    }

    private static async Task<string?> ReadMarkerAsync(
        ExcelMcpService service,
        string sessionId,
        string sheetName = "Sheet1",
        string rangeAddress = "A1")
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.get-values",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                sheetName,
                rangeAddress
            }, ServiceProtocol.JsonOptions)
        });
        RequireSuccess(response);
        using var result = JsonDocument.Parse(response.Result!);
        return result.RootElement.GetProperty("values")[0][0].GetString();
    }

    private static async Task CloseSessionAsync(
        ExcelMcpService service,
        string sessionId,
        bool save,
        WorkbookSaveState? saveState = null)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new { save }, ServiceProtocol.JsonOptions)
        });
        Assert.True(
            response.Success,
            $"{response.ErrorMessage}{Environment.NewLine}" +
            $"HRESULT: {response.HResult ?? "<none>"}{Environment.NewLine}" +
            $"Inner error: {response.InnerError ?? "<none>"}{Environment.NewLine}" +
            $"Pre-save state: {saveState?.ToString() ?? "<not captured>"}");
        Assert.True(string.IsNullOrEmpty(response.ErrorMessage), response.ErrorMessage);
    }

    private static WorkbookSaveState CaptureSaveState(
        ExcelMcpService service,
        string sessionId,
        string workbookPath)
    {
        var batch = Assert.IsAssignableFrom<IExcelBatch>(
            service.SessionManager.GetSession(sessionId));
        var excelState = batch.Execute((context, _) => new
        {
            FullName = context.Book.FullName,
            ReadOnly = context.Book.ReadOnly,
            Saved = context.Book.Saved
        });
        var file = new FileInfo(workbookPath);
        file.Refresh();

        return new WorkbookSaveState(
            workbookPath,
            excelState.FullName,
            excelState.ReadOnly,
            excelState.Saved,
            file.Exists,
            file.Exists && file.IsReadOnly,
            file.Exists ? file.Length : null,
            file.Exists ? file.LastWriteTimeUtc : null,
            batch.ExcelProcessId,
            batch.IsExcelProcessAlive());
    }

    private static string GetSessionId(ServiceResponse response)
    {
        RequireSuccess(response);
        using var result = JsonDocument.Parse(response.Result!);
        var sessionId = result.RootElement.GetProperty("sessionId").GetString();
        Assert.False(string.IsNullOrWhiteSpace(sessionId));
        return sessionId!;
    }

    private static void RequireSuccess(ServiceResponse response)
    {
        Assert.True(response.Success, response.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(response.ErrorMessage), response.ErrorMessage);
    }

    private static async Task AssertSessionIdsAsync(ExcelMcpService service, params string[] expected)
    {
        var response = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
        RequireSuccess(response);
        using var result = JsonDocument.Parse(response.Result!);
        var actual = result.RootElement.GetProperty("sessions").EnumerateArray()
            .Select(item => item.GetProperty("sessionId").GetString()).Order().ToArray();
        Assert.Equal(expected.Order().ToArray(), actual);
        Assert.Equal(expected.Length, service.SessionCount);
    }

    private sealed record WorkflowResult(
        int Index,
        string FilePath,
        string? PersistedValue,
        int ExcelProcessId);

    private sealed record WorkbookSaveState(
        string RequestedPath,
        string ExcelFullName,
        bool ReadOnly,
        bool Saved,
        bool FileExists,
        bool HasReadOnlyAttribute,
        long? FileLength,
        DateTime? LastWriteTimeUtc,
        int? ExcelProcessId,
        bool ProcessAlive);
}
