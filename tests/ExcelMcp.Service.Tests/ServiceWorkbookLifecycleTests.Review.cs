using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    private static readonly string[] ReviewTraceCommands = ["powerquerytrace.list", "powerquerytrace.summarize"];
    private static readonly bool[] ReviewVisibilityStates = [true, false];

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Visibility_DiscoveryStartupCannotUndoNewerOwnedRequest(bool visible)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "visibility-startup.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var name = owned.Execute((context, _) => AddReviewWorkbook(context.App, "startup visibility edits"));
            RequireSuccess(await service.ProcessAsync(new ServiceRequest
            {
                Command = visible ? "window.hide" : "window.show",
                SessionId = ownedId
            }));
            using var read = new ManualResetEventSlim();
            using var release = new ManualResetEventSlim();
            bool? startupVisibility = null;
            ExcelBatch.BeforeStartupVisibilityPublicationHookForTests = batch =>
            {
                if (batch.ProcessIdentity != owned.ProcessIdentity)
                    return;
                batch.BeforeVisibilitySnapshotPublicationHookForTests = () =>
                {
                    startupVisibility = batch.IsExcelVisible;
                    batch.BeforeVisibilitySnapshotPublicationHookForTests = null;
                };
                read.Set();
                Assert.True(release.Wait(TimeSpan.FromSeconds(15)));
            };
            var discovery = Task.Run(() => DiscoverReviewWorkbookAsync(service, name));
            JsonElement workbook;
            try
            {
                Assert.True(read.Wait(TimeSpan.FromSeconds(10)));
                RequireSuccess(await service.ProcessAsync(new ServiceRequest
                {
                    Command = visible ? "window.show" : "window.hide",
                    SessionId = ownedId
                }));
            }
            finally
            {
                release.Set();
                try
                {
                    workbook = await discovery;
                }
                finally
                {
                    ExcelBatch.BeforeStartupVisibilityPublicationHookForTests = null;
                }
            }
            var attachedId = await AttachReviewWorkbookAsync(service, workbook);
            sessions[attachedId] = 0;
            Assert.Equal(visible, startupVisibility);
            Assert.Equal(visible, owned.IsExcelVisible);
            var response = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
            RequireSuccess(response);
            using var result = JsonDocument.Parse(response.Result!);
            Assert.All(result.RootElement.GetProperty("sessions").EnumerateArray(),
                session => Assert.Equal(visible, session.GetProperty("isExcelVisible").GetBoolean()));
            var identity = Assert.IsType<ExcelProcessIdentity>(owned.ProcessIdentity);
            var inspection = ExcelDialogProbe.Inspect(identity, owned.Logger);
            Assert.NotEqual(WorkbookRefreshState.Unknown, inspection.State);
            var windows = inspection.Windows.Where(window =>
                window.ProcessId == identity.ProcessId && window.ClassName == "XLMAIN").ToArray();
            Assert.NotEmpty(windows);
            Assert.Equal(visible, windows.Any(window => window.Visible));
            Assert.Equal("startup visibility edits", await ReadMarkerAsync(service, attachedId));
            var attached = Assert.IsType<AttachedWorkbookBatch>(service.SessionManager.GetSession(attachedId));
            Assert.False(attached.Execute((context, _) => context.Book.Saved));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Visibility_DeferredRecoveryRevealCannotUndoSiblingHide(bool callbackOnAttachment)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "visibility-recovery.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var name = owned.Execute((context, _) => AddReviewWorkbook(context.App, "recovery visibility edits"));
            var id = await AttachReviewWorkbookAsync(service, await DiscoverReviewWorkbookAsync(service, name));
            sessions[id] = 0;
            var attached = Assert.IsType<AttachedWorkbookBatch>(service.SessionManager.GetSession(id));
            var callbackOwner = callbackOnAttachment ? attached.Owner : owned;
            RequireSuccess(await service.ProcessAsync(new ServiceRequest
            {
                Command = "window.show",
                SessionId = ownedId
            }));
            using var entered = new ManualResetEventSlim();
            using var release = new ManualResetEventSlim();
            using var wait = new CancellationTokenSource();
            var operation = Task.Run(() => callbackOwner.Execute((_, _) =>
            {
                entered.Set();
                Assert.True(release.Wait(TimeSpan.FromSeconds(15)));
            }, wait.Token));
            try
            {
                Assert.True(entered.Wait(TimeSpan.FromSeconds(10)));
                wait.Cancel();
                await Assert.ThrowsAnyAsync<OperationCanceledException>(() => operation);
                Assert.True(callbackOwner.OperationStatus.ExecutionPending);
                Assert.True(owned.IsExcelVisible);
                Assert.True(attached.IsExcelVisible);
                RequireSuccess(await service.ProcessAsync(new ServiceRequest
                {
                    Command = "window.hide",
                    SessionId = callbackOnAttachment ? ownedId : id
                }));
            }
            finally
            {
                release.Set();
                await Assert.ThrowsAnyAsync<OperationCanceledException>(() => operation);
                Assert.True(SpinWait.SpinUntil(
                    () => !callbackOwner.OperationStatus.ExecutionPending, TimeSpan.FromSeconds(10)));
                Assert.True(SpinWait.SpinUntil(
                    () => callbackOwner.GetRefreshState() == WorkbookRefreshState.Ready, TimeSpan.FromSeconds(10)));
            }
            var response = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
            RequireSuccess(response);
            using var result = JsonDocument.Parse(response.Result!);
            Assert.All(result.RootElement.GetProperty("sessions").EnumerateArray(),
                session => Assert.False(session.GetProperty("isExcelVisible").GetBoolean()));
            var identity = Assert.IsType<ExcelProcessIdentity>(owned.ProcessIdentity);
            var inspection = ExcelDialogProbe.Inspect(identity, owned.Logger);
            Assert.NotEqual(WorkbookRefreshState.Unknown, inspection.State);
            Assert.DoesNotContain(inspection.Windows,
                window => window.ProcessId == identity.ProcessId && window.ClassName == "XLMAIN" && window.Visible);
            Assert.Equal("recovery visibility edits", await ReadMarkerAsync(service, id));
            Assert.False(attached.Execute((context, _) => context.Book.Saved));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Visibility_OlderCallbackSnapshotCannotUndoOtherConnectionsHide(bool callbackOnAttachment)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "visibility-overlap.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var name = owned.Execute((context, _) => AddReviewWorkbook(context.App, "visibility overlap edits"));
            var id = await AttachReviewWorkbookAsync(service, await DiscoverReviewWorkbookAsync(service, name));
            sessions[id] = 0;
            var attached = Assert.IsType<AttachedWorkbookBatch>(service.SessionManager.GetSession(id));
            RequireSuccess(await service.ProcessAsync(new ServiceRequest
            {
                Command = "window.show",
                SessionId = ownedId
            }));
            using var read = new ManualResetEventSlim();
            using var release = new ManualResetEventSlim();
            var callbackOwner = callbackOnAttachment ? attached.Owner : owned;
            callbackOwner.BeforeVisibilitySnapshotPublicationHookForTests = () =>
            {
                read.Set();
                Assert.True(release.Wait(TimeSpan.FromSeconds(15)));
            };
            var operation = Task.Run(() => callbackOwner.Execute((_, _) => 0));
            try
            {
                Assert.True(read.Wait(TimeSpan.FromSeconds(10)));
                RequireSuccess(await service.ProcessAsync(new ServiceRequest
                {
                    Command = "window.hide",
                    SessionId = callbackOnAttachment ? ownedId : id
                }));
            }
            finally
            {
                release.Set();
                await operation;
                callbackOwner.BeforeVisibilitySnapshotPublicationHookForTests = null;
            }
            var response = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
            RequireSuccess(response);
            using var result = JsonDocument.Parse(response.Result!);
            Assert.All(result.RootElement.GetProperty("sessions").EnumerateArray(),
                session => Assert.False(session.GetProperty("isExcelVisible").GetBoolean()));
            Assert.False(owned.IsExcelVisible);
            Assert.False(attached.IsExcelVisible);
            var identity = Assert.IsType<ExcelProcessIdentity>(owned.ProcessIdentity);
            var inspection = ExcelDialogProbe.Inspect(identity, owned.Logger);
            Assert.NotEqual(WorkbookRefreshState.Unknown, inspection.State);
            Assert.DoesNotContain(inspection.Windows,
                window => window.ProcessId == identity.ProcessId && window.ClassName == "XLMAIN" && window.Visible);
            Assert.Equal("visibility overlap edits", await ReadMarkerAsync(service, id));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Visibility_SharedOwnedAndAttachedConnections_ReportConfirmedWindowState(bool throughAttachment)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "visibility-owner.xlsx"));
            sessions[ownedId] = 0;
            await WriteMarkerAsync(service, ownedId, "owned visibility edits");
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var name = owned.Execute((context, _) => AddReviewWorkbook(context.App, "attached visibility edits"));
            var attachedId = await AttachReviewWorkbookAsync(service,
                await DiscoverReviewWorkbookAsync(service, name));
            sessions[attachedId] = 0;
            var attached = Assert.IsType<AttachedWorkbookBatch>(service.SessionManager.GetSession(attachedId));
            var unrelatedId = await CreateSessionAsync(service, Path.Join(directory, "visibility-unrelated.xlsx"));
            sessions[unrelatedId] = 0;
            var identity = Assert.IsType<ExcelProcessIdentity>(owned.ProcessIdentity);
            Assert.Equal(identity, attached.Owner.ProcessIdentity);
            Assert.False(owned.IsExcelVisible);
            Assert.False(attached.IsExcelVisible);
            foreach (var visible in ReviewVisibilityStates)
            {
                RequireSuccess(await service.ProcessAsync(new ServiceRequest
                {
                    Command = visible ? "window.show" : "window.hide",
                    SessionId = throughAttachment ? attachedId : ownedId
                }));
                using var entered = new ManualResetEventSlim();
                using var release = new ManualResetEventSlim();
                var pending = Task.Run(() => attached.Execute((_, _) =>
                {
                    entered.Set();
                    Assert.True(release.Wait(TimeSpan.FromSeconds(15)));
                }));
                try
                {
                    Assert.True(entered.Wait(TimeSpan.FromSeconds(10)));
                    var timer = System.Diagnostics.Stopwatch.StartNew();
                    var response = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
                    Assert.True(timer.Elapsed < TimeSpan.FromSeconds(5),
                        "Visibility reporting must not wait for a running Excel callback.");
                    RequireSuccess(response);
                    using var result = JsonDocument.Parse(response.Result!);
                    var listed = result.RootElement.GetProperty("sessions").EnumerateArray().ToArray();
                    foreach (var id in new[] { ownedId, attachedId })
                        Assert.Equal(visible, Assert.Single(listed,
                            item => item.GetProperty("sessionId").GetString() == id)
                            .GetProperty("isExcelVisible").GetBoolean());
                    Assert.False(Assert.Single(listed,
                        item => item.GetProperty("sessionId").GetString() == unrelatedId)
                        .GetProperty("isExcelVisible").GetBoolean());
                    var inspection = ExcelDialogProbe.Inspect(identity, owned.Logger);
                    Assert.NotEqual(WorkbookRefreshState.Unknown, inspection.State);
                    var windows = inspection.Windows.Where(window =>
                        window.ProcessId == identity.ProcessId && window.ClassName == "XLMAIN").ToArray();
                    Assert.NotEmpty(windows);
                    Assert.Equal(visible, windows.Any(window => window.Visible));
                }
                finally
                {
                    release.Set();
                    await pending;
                }
            }
            Assert.Equal("owned visibility edits", await ReadMarkerAsync(service, ownedId));
            Assert.Equal("attached visibility edits", await ReadMarkerAsync(service, attachedId));
            Assert.False(owned.Execute((context, _) => context.Book.Saved));
            Assert.False(attached.Execute((context, _) => context.Book.Saved));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Recovery_AttachedSnapshotThenExcelRename_DoesNotApplyOldLocation(bool reattach)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "live-attached-owner.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var name = owned.Execute((context, _) => AddReviewWorkbook(context.App, "live attached edits"));
            var id = await AttachReviewWorkbookAsync(service, await DiscoverReviewWorkbookAsync(service, name));
            sessions[id] = 0;
            var attached = Assert.IsType<AttachedWorkbookBatch>(service.SessionManager.GetSession(id));
            var original = Path.Join(directory, "live-attached-original.xlsx");
            var target = Path.Join(directory, "live-attached-newer.xlsx");
            attached.Execute((context, _) => context.Book.SaveAs(original));
            var selection = await DiscoverReviewWorkbookAsync(service, Path.GetFileName(original));
            Action rename = () => attached.Owner.Execute((_, _) =>
                attached.Owner.GetAttachedWorkbook(attached.WorkbookId).SaveAs(target));
            if (reattach)
                service.SessionManager.AfterAttachmentRefreshHookForTests = rename;
            else
                service.SessionManager.AfterApplicationDiscoveryHookForTests = _ => rename();
            ServiceResponse response;
            try
            {
                response = await service.ProcessAsync(reattach
                    ? new ServiceRequest
                    {
                        Command = "session.attach",
                        Args = JsonSerializer.Serialize(new
                        {
                            workbookId = selection.GetProperty("workbookId").GetString()
                        }, ServiceProtocol.JsonOptions)
                    }
                    : new ServiceRequest { Command = "session.discover" });
            }
            finally
            {
                service.SessionManager.AfterAttachmentRefreshHookForTests = null;
                service.SessionManager.AfterApplicationDiscoveryHookForTests = null;
            }
            if (reattach)
            {
                Assert.False(response.Success);
                Assert.Contains("location changed", response.ErrorMessage, StringComparison.Ordinal);
            }
            else
            {
                RequireSuccess(response);
                using var result = JsonDocument.Parse(response.Result!);
                Assert.False(result.RootElement.GetProperty("complete").GetBoolean());
                Assert.Contains(result.RootElement.GetProperty("errors").EnumerateArray(),
                    error => error.GetProperty("message").GetString()!.Contains(
                        "location changed", StringComparison.Ordinal));
            }
            Assert.Equal(original, attached.WorkbookPath, ignoreCase: true);
            var fresh = await DiscoverReviewWorkbookAsync(service, Path.GetFileName(target));
            Assert.Equal(id, fresh.GetProperty("sessionId").GetString());
            Assert.Equal(selection.GetProperty("workbookId").GetString(), fresh.GetProperty("workbookId").GetString());
            Assert.Equal(id, await AttachReviewWorkbookAsync(service, fresh));
            Assert.Equal(target, attached.WorkbookPath, ignoreCase: true);
            Assert.Throws<InvalidOperationException>(() =>
                service.SessionManager.ReserveSessionFilePath(ownedId, target));
            attached.Execute((context, _) => Assert.Equal(target, context.Book.FullName, ignoreCase: true));
            Assert.Equal("live attached edits", await ReadMarkerAsync(service, id));
            Assert.Equal(2, service.SessionManager.ActiveSessionCount);
        });
    }

    [Fact]
    public async Task Discovery_OwnedMatchThenExcelRename_DoesNotApplyOldLocation()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var original = Path.Join(directory, "live-owned-original.xlsx");
            var intermediate = Path.Join(directory, "live-owned-intermediate.xlsx");
            var target = Path.Join(directory, "live-owned-newer.xlsx");
            var sessionId = await CreateSessionAsync(service, original);
            sessions[sessionId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(sessionId));
            var selection = await DiscoverReviewWorkbookAsync(service, Path.GetFileName(original));
            await WriteMarkerAsync(service, sessionId, "live rename edits");
            owned.Execute((context, _) => context.Book.SaveAs(intermediate));
            service.SessionManager.AfterOwnedWorkbookMatchHookForTests = () =>
            {
                service.SessionManager.AfterOwnedWorkbookMatchHookForTests = null;
                owned.Execute((context, _) => context.Book.SaveAs(target));
            };
            ServiceResponse discovery;
            try
            {
                discovery = await service.ProcessAsync(new ServiceRequest { Command = "session.discover" });
            }
            finally
            {
                service.SessionManager.AfterOwnedWorkbookMatchHookForTests = null;
            }
            RequireSuccess(discovery);
            using var result = JsonDocument.Parse(discovery.Result!);
            Assert.False(result.RootElement.GetProperty("complete").GetBoolean());
            Assert.Contains(result.RootElement.GetProperty("errors").EnumerateArray(),
                error => error.GetProperty("message").GetString()!.Contains(
                    "location changed", StringComparison.Ordinal));
            Assert.Equal(original, owned.WorkbookPath, ignoreCase: true);
            Assert.True(service.SessionManager.TryGetFilePath(sessionId, out var tracked));
            Assert.Equal(original, tracked, ignoreCase: true);
            owned.Execute((context, _) => Assert.Equal(target, context.Book.FullName, ignoreCase: true));
            Assert.Equal("live rename edits", await ReadMarkerAsync(service, sessionId));
            var fresh = await DiscoverReviewWorkbookAsync(service, Path.GetFileName(target));
            Assert.Equal(sessionId, fresh.GetProperty("sessionId").GetString());
            Assert.Equal(selection.GetProperty("workbookId").GetString(), fresh.GetProperty("workbookId").GetString());
            Assert.Equal(sessionId, await AttachReviewWorkbookAsync(service, fresh));
            Assert.Single(service.SessionManager.GetActiveSessions());
            Assert.Equal(target, owned.WorkbookPath, ignoreCase: true);
        });
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task Recovery_OwnedSnapshotBeforeSaveAs_DoesNotPublishStaleLocation(
        bool reattach, bool duringPublication)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var original = Path.Join(directory, "owned-snapshot-original.xlsx");
            var target = Path.Join(directory, "owned-snapshot-newer.xlsx");
            var sessionId = await CreateSessionAsync(service, original);
            sessions[sessionId] = 0;
            await WriteMarkerAsync(service, sessionId, "owned snapshot edits");
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(sessionId));
            var selection = await DiscoverReviewWorkbookAsync(service, Path.GetFileName(original));
            if (!reattach)
                owned.Execute((context, _) => context.Book.SaveAs(
                    Path.Join(directory, "owned-snapshot-intermediate.xlsx")));
            using var snapshotReady = new ManualResetEventSlim();
            using var releaseSnapshot = new ManualResetEventSlim();
            using var publicationReady = new ManualResetEventSlim();
            using var releasePublication = new ManualResetEventSlim();
            using var admitted = new ManualResetEventSlim();
            using var releaseAdmission = new ManualResetEventSlim();
            Action pause = () =>
            {
                snapshotReady.Set();
                Assert.True(releaseSnapshot.Wait(TimeSpan.FromSeconds(30)));
            };
            if (reattach)
                service.SessionManager.AfterAttachmentRefreshHookForTests = pause;
            else
                service.SessionManager.AfterOwnedWorkbookMatchHookForTests = pause;
            Task<ServiceResponse> StartSaveAs() => Task.Run(() => service.ProcessAsync(new ServiceRequest
            {
                Command = "workbook.save-as",
                SessionId = sessionId,
                Args = JsonSerializer.Serialize(new { targetPath = target }, ServiceProtocol.JsonOptions)
            }));
            Task<ServiceResponse>? saving = null;
            if (!reattach)
            {
                service.BeforeSaveAsReservationHookForTests = () =>
                {
                    admitted.Set();
                    Assert.True(releaseAdmission.Wait(TimeSpan.FromSeconds(30)));
                };
                saving = StartSaveAs();
                Assert.True(admitted.Wait(TimeSpan.FromSeconds(10)));
            }
            var older = Task.Run(() => service.ProcessAsync(reattach
                ? new ServiceRequest
                {
                    Command = "session.attach",
                    Args = JsonSerializer.Serialize(new
                    {
                        workbookId = selection.GetProperty("workbookId").GetString()
                    }, ServiceProtocol.JsonOptions)
                }
                : new ServiceRequest { Command = "session.discover" }));
            ServiceResponse? recovery = null;
            try
            {
                Assert.True(snapshotReady.Wait(TimeSpan.FromSeconds(10)));
                if (duringPublication)
                    service.BeforeSaveAsPublicationHookForTests = () =>
                    {
                        publicationReady.Set();
                        Assert.True(releasePublication.Wait(TimeSpan.FromSeconds(30)));
                    };
                saving ??= StartSaveAs();
                releaseAdmission.Set();
                if (duringPublication)
                    Assert.True(publicationReady.Wait(TimeSpan.FromSeconds(10)));
                else if (!reattach)
                    Assert.True(SpinWait.SpinUntil(
                        () => service.SessionManager.TryGetFilePath(sessionId, out var path)
                            && string.Equals(path, target, StringComparison.OrdinalIgnoreCase),
                        TimeSpan.FromSeconds(10)));
                else
                    RequireSuccess(await saving);
                releaseSnapshot.Set();
                recovery = await older;
            }
            finally
            {
                releaseSnapshot.Set();
                releasePublication.Set();
                releaseAdmission.Set();
                service.SessionManager.AfterAttachmentRefreshHookForTests = null;
                service.SessionManager.AfterOwnedWorkbookMatchHookForTests = null;
                service.BeforeSaveAsPublicationHookForTests = null;
                service.BeforeSaveAsReservationHookForTests = null;
                await older;
                if (saving is not null)
                    RequireSuccess(await saving);
                // Track any incorrectly published connection so a failing regression still cleans up.
                foreach (var session in service.SessionManager.GetActiveSessions())
                    sessions[session.SessionId] = 0;
            }
            if (reattach)
            {
                Assert.False(recovery!.Success);
                Assert.Contains("changed", recovery.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            }
            else
            {
                RequireSuccess(recovery!);
                using var result = JsonDocument.Parse(recovery!.Result!);
                Assert.False(result.RootElement.GetProperty("complete").GetBoolean());
                Assert.Contains(result.RootElement.GetProperty("errors").EnumerateArray(),
                    error => error.GetProperty("message").GetString()!.Contains(
                        "changed", StringComparison.OrdinalIgnoreCase));
            }
            Assert.Equal(sessionId, Assert.Single(service.SessionManager.GetActiveSessions()).SessionId);
            Assert.Equal(target, owned.WorkbookPath, ignoreCase: true);
            Assert.True(service.SessionManager.TryGetFilePath(sessionId, out var tracked));
            Assert.Equal(target, tracked, ignoreCase: true);
            owned.Execute((context, _) => Assert.Equal(target, context.Book.FullName, ignoreCase: true));
            Assert.Equal("owned snapshot edits", await ReadMarkerAsync(service, sessionId));
            var fresh = await DiscoverReviewWorkbookAsync(service, Path.GetFileName(target));
            Assert.Equal(sessionId, fresh.GetProperty("sessionId").GetString());
            Assert.Equal(selection.GetProperty("workbookId").GetString(), fresh.GetProperty("workbookId").GetString());
            Assert.Equal(sessionId, await AttachReviewWorkbookAsync(service, fresh));
            var neighborId = await CreateSessionAsync(service, Path.Join(directory, "owned-snapshot-neighbor.xlsx"));
            sessions[neighborId] = 0;
            Assert.Throws<InvalidOperationException>(() =>
                service.SessionManager.ReserveSessionFilePath(neighborId, target));
        });
    }

    [Fact]
    public async Task Recovery_HealthyAttachmentCannotClearRefreshingNeighborsTimeout()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "timeout-neighbors.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var nameA = owned.Execute((context, _) => AddReviewWorkbook(context.App, "refreshing edits"));
            var nameB = owned.Execute((context, _) => AddReviewWorkbook(context.App, "healthy edits"));
            var idA = await AttachReviewWorkbookAsync(service, await DiscoverReviewWorkbookAsync(service, nameA));
            sessions[idA] = 0;
            var idB = await AttachReviewWorkbookAsync(service, await DiscoverReviewWorkbookAsync(service, nameB));
            sessions[idB] = 0;
            var a = Assert.IsType<AttachedWorkbookBatch>(service.SessionManager.GetSession(idA));
            var b = Assert.IsType<AttachedWorkbookBatch>(service.SessionManager.GetSession(idB));
            Assert.Same(a.Owner, b.Owner);
            RequireSuccess(await service.ProcessAsync(new ServiceRequest
            {
                Command = "powerquery.create",
                SessionId = idA,
                Args = """{"queryName":"SharedRecovery","mCode":"#table({\"Marker\"}, {{\"Before\"}})","loadDestination":"worksheet","targetSheet":"SharedRecovery"}"""
            }));
            RequireSuccess(await service.ProcessAsync(new ServiceRequest
            {
                Command = "powerquery.update",
                SessionId = idA,
                Args = """{"queryName":"SharedRecovery","mCode":"Function.InvokeAfter(() => #table({\"Marker\"}, {{\"After\"}}), #duration(0,0,0,12))","refresh":false}"""
            }));
            using var started = new ManualResetEventSlim();
            using var release = new ManualResetEventSlim();
            using var wait = new CancellationTokenSource();
            var operation = Task.Run(() => a.Execute((context, _) =>
            {
                Excel.Sheets? sheets = null;
                Excel.Worksheet? sheet = null;
                Excel.ListObjects? tables = null;
                Excel.ListObject? table = null;
                Excel.QueryTable? query = null;
                try
                {
                    sheets = context.Book.Worksheets;
                    sheet = (Excel.Worksheet)sheets["SharedRecovery"];
                    tables = sheet.ListObjects;
                    table = tables[1];
                    query = table.QueryTable;
                    Assert.True(query.Refresh(true));
                    Assert.True(query.Refreshing);
                    started.Set();
                    Assert.True(release.Wait(TimeSpan.FromSeconds(15)));
                }
                finally
                {
                    ComUtilities.Release(ref query);
                    ComUtilities.Release(ref table);
                    ComUtilities.Release(ref tables);
                    ComUtilities.Release(ref sheet);
                    ComUtilities.Release(ref sheets);
                }
            }, wait.Token));
            var called = false;
            try
            {
                Assert.True(started.Wait(TimeSpan.FromSeconds(10)));
                wait.Cancel();
                await Assert.ThrowsAnyAsync<OperationCanceledException>(() => operation);
                release.Set();
                Assert.True(SpinWait.SpinUntil(
                    () => !a.Owner.OperationStatus.ExecutionPending, TimeSpan.FromSeconds(10)));
                Assert.True(a.HasTimedOutOperation);
                var state = b.GetRefreshState();
                var deadline = DateTime.UtcNow.AddSeconds(3);
                while (state == WorkbookRefreshState.Unknown && DateTime.UtcNow < deadline)
                {
                    await Task.Delay(50);
                    state = b.GetRefreshState();
                }
                Assert.Equal(WorkbookRefreshState.Refreshing, state);
                Assert.True(a.HasTimedOutOperation);
                Assert.Throws<TimeoutException>(() => a.Execute((_, _) => called = true));
                Assert.False(called);
            }
            finally
            {
                release.Set();
                await Assert.ThrowsAnyAsync<OperationCanceledException>(() => operation);
                Assert.True(SpinWait.SpinUntil(
                    () => a.Owner.GetRefreshState() == WorkbookRefreshState.Ready,
                    TimeSpan.FromSeconds(45)));
            }
            Assert.False(a.HasTimedOutOperation);
            Assert.Equal("After", await ReadMarkerAsync(service, idA, "SharedRecovery", "A2"));
            Assert.Equal("refreshing edits", await ReadMarkerAsync(service, idA));
            Assert.Equal("healthy edits", await ReadMarkerAsync(service, idB));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Discovery_OrdinaryCommandAfterExcelSaveAs_PreservesNewerLocation(bool rediscover)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "context-owner.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var name = owned.Execute((context, _) => AddReviewWorkbook(context.App, "context rename edits"));
            var attachedId = await AttachReviewWorkbookAsync(service,
                await DiscoverReviewWorkbookAsync(service, name));
            sessions[attachedId] = 0;
            var attached = Assert.IsType<AttachedWorkbookBatch>(service.SessionManager.GetSession(attachedId));
            var original = Path.Join(directory, "context-original.xlsx");
            var target = Path.Join(directory, "context-newer.xlsx");
            attached.Execute((context, _) => context.Book.SaveAs(original));
            await DiscoverReviewWorkbookAsync(service, "context-original.xlsx");
            using var scanned = new ManualResetEventSlim();
            using var release = new ManualResetEventSlim();
            var scans = 0;
            service.SessionManager.AfterApplicationDiscoveryHookForTests = application =>
            {
                if (!ReferenceEquals(application, attached.Owner) || Interlocked.Increment(ref scans) != 1)
                    return;
                scanned.Set();
                Assert.True(release.Wait(TimeSpan.FromSeconds(30)));
            };
            var older = Task.Run(() => service.ProcessAsync(new ServiceRequest { Command = "session.discover" }));
            ServiceResponse? response = null;
            try
            {
                Assert.True(scanned.Wait(TimeSpan.FromSeconds(10)));
                attached.Owner.Execute((_, _) => attached.Owner.GetAttachedWorkbook(attached.WorkbookId).SaveAs(target));
                Assert.Equal("context rename edits", await ReadMarkerAsync(service, attachedId));
                try
                {
                    Assert.Throws<InvalidOperationException>(() =>
                        service.SessionManager.ReserveSessionFilePath(ownedId, target));
                }
                finally
                {
                    service.SessionManager.ReleaseSessionFilePathReservation(ownedId, target);
                }
                if (rediscover)
                    await DiscoverReviewWorkbookAsync(service, "context-newer.xlsx");
            }
            finally
            {
                release.Set();
                service.SessionManager.AfterApplicationDiscoveryHookForTests = null;
                response = await older;
            }
            RequireSuccess(response);
            using var result = JsonDocument.Parse(response.Result!);
            Assert.Equal(target, attached.WorkbookPath, ignoreCase: true);
            Assert.False(result.RootElement.GetProperty("complete").GetBoolean());
            Assert.Contains(result.RootElement.GetProperty("errors").EnumerateArray(),
                error => error.GetProperty("message").GetString()!.Contains(
                    "changed during discovery", StringComparison.Ordinal));
            await DiscoverReviewWorkbookAsync(service, "context-newer.xlsx");
            var listed = Assert.Single(service.SessionManager.GetActiveSessions(), item => item.SessionId == attachedId);
            Assert.Equal(target, listed.FilePath, ignoreCase: true);
            Assert.Equal("context-newer.xlsx", listed.WorkbookName);
            Assert.Throws<InvalidOperationException>(() => service.SessionManager.ReserveSessionFilePath(ownedId, target));
            attached.Execute((context, _) => Assert.Equal(target, context.Book.FullName, ignoreCase: true));
            Assert.Equal("context rename edits", await ReadMarkerAsync(service, attachedId));
        });
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task Discovery_CleanupFailsThenProcessExits_ReleasesConnection(bool rediscover, bool retainOtherLease)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "discovery-cleanup.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var identity = Assert.IsType<ExcelProcessIdentity>(owned.ProcessIdentity);
            var name = owned.Execute((context, _) => AddReviewWorkbook(context.App, "cleanup neighbor"));
            var attachedId = await AttachReviewWorkbookAsync(service,
                await DiscoverReviewWorkbookAsync(service, name));
            sessions[attachedId] = 0;
            var attached = Assert.IsType<AttachedWorkbookBatch>(service.SessionManager.GetSession(attachedId));
            await CloseSessionAsync(service, ownedId, save: false);
            sessions.TryRemove(ownedId, out _);
            await CloseSessionAsync(service, attachedId, save: false);
            sessions.TryRemove(attachedId, out _);
            using var process = System.Diagnostics.Process.GetProcessById(identity.ProcessId);
            using var otherLease = retainOtherLease ? new ExcelApplicationLease(identity) : null;
            attached.Owner.BeforeApplicationLeaseReleaseHookForTests = () =>
                throw new InvalidOperationException("Test-only discovery connection cleanup failure.");
            var failure = Assert.Throws<AggregateException>(() => service.SessionManager.Dispose());
            Assert.Contains("Test-only discovery connection cleanup failure.", failure.ToString(), StringComparison.Ordinal);
            await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(15));
            Assert.True(OwnedProcessGuard.TryConfirmExited(identity));
            Assert.True(service.SessionManager.HasActiveConnections);
            if (rediscover)
            {
                await Assert.ThrowsAsync<ObjectDisposedException>(
                    () => service.SessionManager.DiscoverOpenWorkbooksAsync());
                Assert.True(service.SessionManager.HasActiveConnections);
            }
            service.SessionManager.Dispose();
            Assert.False(service.SessionManager.HasActiveConnections);
            Assert.Equal(!retainOtherLease, await CanOtherHostAcquireReviewLeaseAsync(identity));
            otherLease?.Dispose();
            Assert.True(await CanOtherHostAcquireReviewLeaseAsync(identity));
        });
    }

    [Theory]
    [InlineData(".xlsx", true)]
    [InlineData(".xlsm", false)]
    public async Task SaveCopyAs_NeverSavedAttachment_PreservesSource(string extension, bool expectedSuccess)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "copy-owner.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var name = owned.Execute((context, _) => AddReviewWorkbook(context.App, "never-saved copy"));
            var selection = await DiscoverReviewWorkbookAsync(service, name);
            var attachedId = await AttachReviewWorkbookAsync(service, selection);
            sessions[attachedId] = 0;
            var attached = Assert.IsType<AttachedWorkbookBatch>(service.SessionManager.GetSession(attachedId));
            var targetPath = Path.Join(directory, "copy" + extension);
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "workbook.save-copy-as",
                SessionId = attachedId,
                Args = JsonSerializer.Serialize(new { targetPath }, ServiceProtocol.JsonOptions)
            });
            Assert.Equal(expectedSuccess, response.Success);
            if (expectedSuccess)
            {
                RequireSuccess(response);
                using var result = JsonDocument.Parse(response.Result!);
                Assert.Equal(targetPath, result.RootElement.GetProperty("filePath").GetString(), ignoreCase: true);
                var copyId = await OpenSessionAsync(service, targetPath);
                sessions[copyId] = 0;
                Assert.Equal("never-saved copy", await ReadMarkerAsync(service, copyId));
                await CloseSessionAsync(service, copyId, save: false);
                sessions.TryRemove(copyId, out _);
            }
            else
            {
                Assert.Contains("'.xlsx'", response.ErrorMessage, StringComparison.Ordinal);
                Assert.False(File.Exists(targetPath));
            }
            attached.Execute((context, _) =>
            {
                Assert.Equal(name, context.Book.Name);
                Assert.Equal(name, context.Book.FullName);
                Assert.Equal(string.Empty, context.Book.Path);
                Assert.False(context.Book.Saved);
            });
            var listed = Assert.Single(service.SessionManager.GetActiveSessions(), item => item.SessionId == attachedId);
            Assert.Null(listed.FilePath);
            Assert.Equal(name, listed.WorkbookName);
            Assert.Equal("never-saved copy", await ReadMarkerAsync(service, attachedId));
        });
    }

    [Fact]
    public async Task Close_AttachedWorkbookWithSlowReadiness_VerifiesStateOnExcelThread()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "slow-attached-owner.xlsx"));
            sessions[ownedId] = 0;
            await WriteMarkerAsync(service, ownedId, "neighbor unsaved");
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var name = owned.Execute((context, _) => AddReviewWorkbook(context.App, "selected unsaved"));
            var attachedId = await AttachReviewWorkbookAsync(service,
                await DiscoverReviewWorkbookAsync(service, name));
            sessions[attachedId] = 0;
            var attached = Assert.IsType<AttachedWorkbookBatch>(service.SessionManager.GetSession(attachedId));
            attached.Owner.BeforeRefreshStateReadHookForTests = () => Thread.Sleep(TimeSpan.FromSeconds(1.5));
            try
            {
                Assert.Equal(WorkbookRefreshState.Unknown, attached.GetRefreshState());
                Assert.True(SpinWait.SpinUntil(
                    () => service.SessionManager.ValidateClose(attachedId).CanClose,
                    TimeSpan.FromSeconds(10)));
                await CloseSessionAsync(service, attachedId, save: false);
                sessions.TryRemove(attachedId, out _);
            }
            finally
            {
                attached.Owner.BeforeRefreshStateReadHookForTests = null;
                attached.Owner.Execute((_, _) => 0);
            }
            Assert.Equal(ownedId, Assert.Single(service.SessionManager.GetActiveSessions()).SessionId);
            Assert.Equal("neighbor unsaved", await ReadMarkerAsync(service, ownedId));
            owned.Execute((context, _) =>
            {
                Excel.Workbooks? workbooks = null;
                try
                {
                    workbooks = context.App.Workbooks;
                    Assert.Equal(1, workbooks.Count);
                    Assert.False(context.Book.Saved);
                }
                finally
                {
                    ComUtilities.Release(ref workbooks);
                }
            });
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Save_AttachedWorkbookWithSlowReadiness_PersistsSelectedWorkbook(bool close)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "slow-save-owner.xlsx"));
            sessions[ownedId] = 0;
            await WriteMarkerAsync(service, ownedId, "neighbor unsaved");
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var name = owned.Execute((context, _) => AddReviewWorkbook(context.App, "initial selected value"));
            var attachedId = await AttachReviewWorkbookAsync(service,
                await DiscoverReviewWorkbookAsync(service, name));
            sessions[attachedId] = 0;
            var path = Path.Join(directory, "slow-save-selected.xlsx");
            RequireSuccess(await service.ProcessAsync(new ServiceRequest
            {
                Command = "workbook.save-as",
                SessionId = attachedId,
                Args = JsonSerializer.Serialize(new { targetPath = path }, ServiceProtocol.JsonOptions)
            }));
            await WriteMarkerAsync(service, attachedId, "persist selected edit");
            var attached = Assert.IsType<AttachedWorkbookBatch>(service.SessionManager.GetSession(attachedId));
            attached.Owner.BeforeRefreshStateReadHookForTests = () => Thread.Sleep(TimeSpan.FromSeconds(1.5));
            try
            {
                Assert.Equal(WorkbookRefreshState.Unknown, attached.GetRefreshState());
                Assert.True(SpinWait.SpinUntil(
                    () => service.SessionManager.ValidateClose(attachedId).CanClose,
                    TimeSpan.FromSeconds(10)));
                RequireSuccess(await service.ProcessAsync(new ServiceRequest
                {
                    Command = close ? "session.close" : "session.save",
                    SessionId = attachedId,
                    Args = close ? """{"save":true}""" : null
                }));
                if (close)
                    sessions.TryRemove(attachedId, out _);
            }
            finally
            {
                attached.Owner.BeforeRefreshStateReadHookForTests = null;
                attached.Owner.Execute((_, _) => 0);
            }
            if (!close)
            {
                await CloseSessionAsync(service, attachedId, save: false);
                sessions.TryRemove(attachedId, out _);
            }
            var reopenedId = await OpenSessionAsync(service, path);
            sessions[reopenedId] = 0;
            Assert.Equal("persist selected edit", await ReadMarkerAsync(service, reopenedId));
            Assert.Equal("neighbor unsaved", await ReadMarkerAsync(service, ownedId));
            Assert.False(owned.Execute((context, _) => context.Book.Saved));
        });
    }

    [Theory]
    [InlineData(false, true, false)]
    [InlineData(true, true, false)]
    [InlineData(true, false, false)]
    [InlineData(true, false, true)]
    public async Task Recovery_OlderSnapshot_PreservesSaveAsPublication(
        bool reattach, bool duringPublication, bool returnToOriginal)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "publication-owner.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var name = owned.Execute((context, _) => AddReviewWorkbook(context.App, "preserve publication edits"));
            var selection = await DiscoverReviewWorkbookAsync(service, name);
            var attachedId = await AttachReviewWorkbookAsync(service, selection);
            sessions[attachedId] = 0;
            var attached = Assert.IsType<AttachedWorkbookBatch>(service.SessionManager.GetSession(attachedId));
            var targetPath = Path.Join(directory, "publication-target.xlsx");
            if (returnToOriginal)
                RequireSuccess(await service.ProcessAsync(new ServiceRequest
                {
                    Command = "workbook.save-as",
                    SessionId = attachedId,
                    Args = JsonSerializer.Serialize(new { targetPath }, ServiceProtocol.JsonOptions)
                }));
            using var snapshotReady = new ManualResetEventSlim();
            using var releaseSnapshot = new ManualResetEventSlim();
            using var publicationReady = new ManualResetEventSlim();
            using var releasePublication = new ManualResetEventSlim();
            Action pauseSnapshot = () =>
            {
                snapshotReady.Set();
                Assert.True(releaseSnapshot.Wait(TimeSpan.FromSeconds(30)));
            };
            if (reattach)
                service.SessionManager.AfterAttachmentRefreshHookForTests = pauseSnapshot;
            else
                service.SessionManager.AfterApplicationDiscoveryHookForTests = application =>
                {
                    if (ReferenceEquals(application, attached.Owner))
                        pauseSnapshot();
                };
            var older = Task.Run(() => service.ProcessAsync(reattach
                ? new ServiceRequest
                {
                    Command = "session.attach",
                    Args = JsonSerializer.Serialize(new
                    {
                        workbookId = selection.GetProperty("workbookId").GetString()
                    }, ServiceProtocol.JsonOptions)
                }
                : new ServiceRequest { Command = "session.discover" }));
            Task<ServiceResponse>? saving = null;
            ServiceResponse? recovery = null;
            try
            {
                Assert.True(snapshotReady.Wait(TimeSpan.FromSeconds(10)));
                if (returnToOriginal)
                    RequireSuccess(await service.ProcessAsync(new ServiceRequest
                    {
                        Command = "workbook.save-as",
                        SessionId = attachedId,
                        Args = JsonSerializer.Serialize(new
                        {
                            targetPath = Path.Join(directory, "intermediate-target.xlsx")
                        }, ServiceProtocol.JsonOptions)
                    }));
                if (duringPublication)
                    service.BeforeSaveAsPublicationHookForTests = () =>
                    {
                        publicationReady.Set();
                        Assert.True(releasePublication.Wait(TimeSpan.FromSeconds(30)));
                    };
                saving = Task.Run(() => service.ProcessAsync(new ServiceRequest
                {
                    Command = "workbook.save-as",
                    SessionId = attachedId,
                    Args = JsonSerializer.Serialize(new
                    {
                        targetPath,
                        overwrite = returnToOriginal
                    }, ServiceProtocol.JsonOptions)
                }));
                if (duringPublication)
                    Assert.True(publicationReady.Wait(TimeSpan.FromSeconds(10)));
                else
                    RequireSuccess(await saving);
                releaseSnapshot.Set();
                recovery = await older;
                Assert.Equal(targetPath, attached.WorkbookPath, ignoreCase: true);
            }
            finally
            {
                releaseSnapshot.Set();
                releasePublication.Set();
                service.BeforeSaveAsPublicationHookForTests = null;
                service.SessionManager.AfterAttachmentRefreshHookForTests = null;
                service.SessionManager.AfterApplicationDiscoveryHookForTests = null;
                await older;
                if (saving is not null)
                    RequireSuccess(await saving);
            }
            if (reattach)
            {
                Assert.False(recovery!.Success);
                Assert.Contains("changed", recovery.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            }
            else
            {
                RequireSuccess(recovery!);
                using var result = JsonDocument.Parse(recovery!.Result!);
                Assert.False(result.RootElement.GetProperty("complete").GetBoolean());
                Assert.Contains(result.RootElement.GetProperty("errors").EnumerateArray(),
                    error => error.GetProperty("message").GetString()!.Contains(
                        "Save As", StringComparison.Ordinal));
            }
            var listed = Assert.Single(service.SessionManager.GetActiveSessions(),
                session => session.SessionId == attachedId);
            Assert.Equal(targetPath, listed.FilePath, ignoreCase: true);
            Assert.Equal("publication-target.xlsx", listed.WorkbookName);
            Assert.Equal("preserve publication edits", await ReadMarkerAsync(service, attachedId));
            attached.Execute((context, _) => Assert.Equal(targetPath, context.Book.FullName, ignoreCase: true));
            Assert.Equal(attachedId, await AttachReviewWorkbookAsync(service, selection));
            Assert.Throws<InvalidOperationException>(() =>
                service.SessionManager.ReserveSessionFilePath(ownedId, targetPath));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Attach_WorkbookClosesAfterRefresh_DoesNotPublishStaleSession(bool initiallyAttached)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "attach-close-owner.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var name = owned.Execute((context, _) => AddReviewWorkbook(context.App, "close selected only"));
            var selection = await DiscoverReviewWorkbookAsync(service, name);
            string? attachedId = null;
            if (initiallyAttached)
            {
                attachedId = await AttachReviewWorkbookAsync(service, selection);
                sessions[attachedId] = 0;
            }
            using var refreshed = new ManualResetEventSlim();
            using var release = new ManualResetEventSlim();
            var refreshes = 0;
            service.SessionManager.AfterAttachmentRefreshHookForTests = () =>
            {
                if (Interlocked.Increment(ref refreshes) != 1)
                    return;
                refreshed.Set();
                Assert.True(release.Wait(TimeSpan.FromSeconds(30)));
            };
            var older = Task.Run(() => service.ProcessAsync(new ServiceRequest
            {
                Command = "session.attach",
                Args = JsonSerializer.Serialize(new
                {
                    workbookId = selection.GetProperty("workbookId").GetString()
                }, ServiceProtocol.JsonOptions)
            }));
            ServiceResponse? response = null;
            try
            {
                Assert.True(refreshed.Wait(TimeSpan.FromSeconds(10)));
                if (attachedId is null)
                {
                    attachedId = await AttachReviewWorkbookAsync(service, selection);
                    sessions[attachedId] = 0;
                }
                RequireSuccess(await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.close",
                    SessionId = attachedId,
                    Args = """{"save":false}"""
                }));
                sessions.TryRemove(attachedId, out _);
            }
            finally
            {
                release.Set();
                service.SessionManager.AfterAttachmentRefreshHookForTests = null;
                response = await older;
                if (response.Success)
                {
                    using var result = JsonDocument.Parse(response.Result!);
                    sessions[result.RootElement.GetProperty("sessionId").GetString()!] = 0;
                }
            }
            Assert.False(response.Success);
            Assert.Contains("changed", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(ownedId, Assert.Single(service.SessionManager.GetActiveSessions()).SessionId);
            owned.Execute((context, _) =>
            {
                Excel.Workbooks? workbooks = null;
                try
                {
                    workbooks = context.App.Workbooks;
                    Assert.Equal(1, workbooks.Count);
                    Assert.NotEqual(name, context.Book.Name);
                }
                finally
                {
                    ComUtilities.Release(ref workbooks);
                }
            });
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Attach_OverlappingFirstConnection_PreservesNewerSession(bool saveAs)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "first-attach-owner.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var name = owned.Execute((context, _) => AddReviewWorkbook(context.App, "first attachment edits"));
            var selection = await DiscoverReviewWorkbookAsync(service, name);
            using var refreshed = new ManualResetEventSlim();
            using var release = new ManualResetEventSlim();
            var refreshes = 0;
            service.SessionManager.AfterAttachmentRefreshHookForTests = () =>
            {
                if (Interlocked.Increment(ref refreshes) != 1)
                    return;
                refreshed.Set();
                Assert.True(release.Wait(TimeSpan.FromSeconds(30)));
            };
            var older = Task.Run(() => service.ProcessAsync(new ServiceRequest
            {
                Command = "session.attach",
                Args = JsonSerializer.Serialize(new
                {
                    workbookId = selection.GetProperty("workbookId").GetString()
                }, ServiceProtocol.JsonOptions)
            }));
            ServiceResponse? response = null;
            string? attachedId = null;
            var targetPath = Path.Join(directory, "newly-attached.xlsx");
            try
            {
                Assert.True(refreshed.Wait(TimeSpan.FromSeconds(10)));
                attachedId = await AttachReviewWorkbookAsync(service, selection);
                sessions[attachedId] = 0;
                if (saveAs)
                    RequireSuccess(await service.ProcessAsync(new ServiceRequest
                    {
                        Command = "workbook.save-as",
                        SessionId = attachedId,
                        Args = JsonSerializer.Serialize(new { targetPath }, ServiceProtocol.JsonOptions)
                    }));
            }
            finally
            {
                release.Set();
                service.SessionManager.AfterAttachmentRefreshHookForTests = null;
                response = await older;
            }
            if (saveAs)
            {
                Assert.False(response.Success);
                Assert.Contains("changed", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            }
            else
            {
                RequireSuccess(response);
                using var result = JsonDocument.Parse(response.Result!);
                Assert.Equal(attachedId, result.RootElement.GetProperty("sessionId").GetString());
            }
            var attached = Assert.IsType<AttachedWorkbookBatch>(service.SessionManager.GetSession(attachedId!));
            Assert.Equal(saveAs ? targetPath : name, attached.WorkbookPath, ignoreCase: true);
            Assert.Equal("first attachment edits", await ReadMarkerAsync(service, attachedId!));
            attached.Execute((context, _) =>
            {
                Assert.Equal(saveAs ? targetPath : name, context.Book.FullName, ignoreCase: true);
                Assert.Equal(saveAs, context.Book.Saved);
            });
            Assert.Equal(2, service.SessionManager.GetActiveSessions().Count);
        });
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task Discovery_OlderSnapshot_DoesNotUndoNewerSaveAs(bool ownedClose, bool rediscover)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "metadata-owner.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var name = owned.Execute((context, _) => AddReviewWorkbook(context.App, "keep renamed value"));
            var attachedId = await AttachReviewWorkbookAsync(service,
                await DiscoverReviewWorkbookAsync(service, name));
            sessions[attachedId] = 0;
            var attached = Assert.IsType<AttachedWorkbookBatch>(service.SessionManager.GetSession(attachedId));
            using var enumerated = new ManualResetEventSlim();
            using var release = new ManualResetEventSlim();
            var scans = 0;
            Action<ExcelBatch> pause = application =>
            {
                if (!ReferenceEquals(application, attached.Owner)
                    || Interlocked.Increment(ref scans) != 1)
                    return;
                enumerated.Set();
                Assert.True(release.Wait(TimeSpan.FromSeconds(30)));
            };
            if (ownedClose)
                service.SessionManager.AfterOwnedCloseDiscoveryHookForTests = pause;
            else
                service.SessionManager.AfterApplicationDiscoveryHookForTests = pause;
            var older = Task.Run(() => service.ProcessAsync(ownedClose
                ? new ServiceRequest { Command = "session.close", SessionId = ownedId, Args = """{"save":false}""" }
                : new ServiceRequest { Command = "session.discover" }));
            ServiceResponse? olderResponse = null;
            var targetPath = Path.Join(directory, "newer-attached.xlsx");
            try
            {
                Assert.True(enumerated.Wait(TimeSpan.FromSeconds(10)));
                RequireSuccess(await service.ProcessAsync(new ServiceRequest
                {
                    Command = "workbook.save-as",
                    SessionId = attachedId,
                    Args = JsonSerializer.Serialize(new { targetPath }, ServiceProtocol.JsonOptions)
                }));
                if (rediscover)
                    await DiscoverReviewWorkbookAsync(service, "newer-attached.xlsx");
            }
            finally
            {
                release.Set();
                service.SessionManager.AfterApplicationDiscoveryHookForTests = null;
                service.SessionManager.AfterOwnedCloseDiscoveryHookForTests = null;
                olderResponse = await older;
                RequireSuccess(olderResponse);
                if (ownedClose)
                    sessions.TryRemove(ownedId, out _);
            }
            if (!ownedClose)
            {
                using var result = JsonDocument.Parse(olderResponse!.Result!);
                Assert.False(result.RootElement.GetProperty("complete").GetBoolean());
                Assert.Contains(result.RootElement.GetProperty("errors").EnumerateArray(),
                    error => error.GetProperty("message").GetString()!.Contains(
                        "changed during discovery", StringComparison.Ordinal));
            }
            var listed = Assert.Single(service.SessionManager.GetActiveSessions(),
                session => session.SessionId == attachedId);
            Assert.Equal(targetPath, listed.FilePath, ignoreCase: true);
            Assert.Equal("newer-attached.xlsx", listed.WorkbookName);
            Assert.Equal(targetPath, attached.WorkbookPath, ignoreCase: true);
            var probeId = await CreateSessionAsync(service, Path.Join(directory, "reservation-probe.xlsx"));
            sessions[probeId] = 0;
            Assert.Throws<InvalidOperationException>(() =>
                service.SessionManager.ReserveSessionFilePath(probeId, targetPath));
            Assert.Equal("keep renamed value", await ReadMarkerAsync(service, attachedId));
        });
    }

    [Fact]
    public async Task SaveAs_AttachedWorkbook_ImmediatelyReportsNewNameAndPath()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "save-as-owner.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var name = owned.Execute((context, _) => AddReviewWorkbook(context.App, "attached rename"));
            var attachedId = await AttachReviewWorkbookAsync(service,
                await DiscoverReviewWorkbookAsync(service, name));
            sessions[attachedId] = 0;
            var targetPath = Path.Join(directory, "renamed-attached.xlsx");
            RequireSuccess(await service.ProcessAsync(new ServiceRequest
            {
                Command = "workbook.save-as",
                SessionId = attachedId,
                Args = JsonSerializer.Serialize(new { targetPath }, ServiceProtocol.JsonOptions)
            }));

            var listed = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
            RequireSuccess(listed);
            using var result = JsonDocument.Parse(listed.Result!);
            var session = Assert.Single(result.RootElement.GetProperty("sessions").EnumerateArray(),
                item => item.GetProperty("sessionId").GetString() == attachedId);
            Assert.Equal(targetPath, session.GetProperty("filePath").GetString(), ignoreCase: true);
            Assert.Equal("renamed-attached.xlsx", session.GetProperty("workbookName").GetString());
            Assert.Equal("attached rename", await ReadMarkerAsync(service, attachedId));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SaveAs_AfterAttachedLocationConflict_PreservesOtherSessionsReservation(bool releaseCompetingReservation)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "reservation-owner.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var name = owned.Execute((context, _) => AddReviewWorkbook(context.App, "keep conflict edits"));
            var attachedId = await AttachReviewWorkbookAsync(service,
                await DiscoverReviewWorkbookAsync(service, name));
            sessions[attachedId] = 0;
            var attached = Assert.IsType<AttachedWorkbookBatch>(service.SessionManager.GetSession(attachedId));
            var conflictPath = Path.Join(directory, "reserved-target.xlsx");
            var competingName = owned.Execute((context, _) => AddReviewWorkbook(context.App, "keep competing edits"));
            var competingId = await AttachReviewWorkbookAsync(service,
                await DiscoverReviewWorkbookAsync(service, competingName));
            sessions[competingId] = 0;
            var reservation = service.SessionManager.ReserveSessionFilePath(competingId, conflictPath);
            try
            {
                attached.Execute((context, _) => context.Book.SaveAs(conflictPath));
                var discovery = await service.SessionManager.DiscoverOpenWorkbooksAsync();
                Assert.False(discovery.Complete);
                Assert.Contains(discovery.Errors, error =>
                    error.Message.Contains($"reserved by session '{competingId}'", StringComparison.Ordinal));
                Assert.Throws<InvalidOperationException>(() =>
                    service.SessionManager.ReserveSessionFilePath(attachedId, conflictPath));
                if (releaseCompetingReservation)
                {
                    service.SessionManager.ReleaseSessionFilePathReservation(competingId, reservation);
                    try
                    {
                        Assert.Throws<InvalidOperationException>(() =>
                            service.SessionManager.ReserveSessionFilePath(competingId, conflictPath));
                        var open = await service.ProcessAsync(new ServiceRequest
                        {
                            Command = "session.open",
                            Args = JsonSerializer.Serialize(new { filePath = conflictPath }, ServiceProtocol.JsonOptions)
                        });
                        Assert.False(open.Success);
                        Assert.Contains("already open", open.ErrorMessage, StringComparison.OrdinalIgnoreCase);
                        Assert.Equal(conflictPath, attached.WorkbookPath, ignoreCase: true);
                        Assert.Equal("keep conflict edits", await ReadMarkerAsync(service, attachedId));
                        Assert.Equal("keep competing edits", await ReadMarkerAsync(service, competingId));
                        await CloseSessionAsync(service, attachedId, save: false);
                        sessions.TryRemove(attachedId, out _);
                        Assert.Equal(conflictPath,
                            service.SessionManager.ReserveSessionFilePath(competingId, conflictPath),
                            ignoreCase: true);
                    }
                    finally
                    {
                        service.SessionManager.ReleaseSessionFilePathReservation(competingId, conflictPath);
                    }
                    return;
                }
                var targetPath = Path.Join(directory, "conflict-resolved.xlsx");
                RequireSuccess(await service.ProcessAsync(new ServiceRequest
                {
                    Command = "workbook.save-as",
                    SessionId = attachedId,
                    Args = JsonSerializer.Serialize(new { targetPath }, ServiceProtocol.JsonOptions)
                }));
                Assert.Equal("keep conflict edits", await ReadMarkerAsync(service, attachedId));
                Assert.Equal(targetPath, attached.WorkbookPath, ignoreCase: true);
                Assert.Throws<InvalidOperationException>(() =>
                    service.SessionManager.ReserveSessionFilePath(attachedId, conflictPath));
            }
            finally
            {
                service.SessionManager.ReleaseSessionFilePathReservation(competingId, reservation);
            }
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Close_CleanupFailsThenExactProcessExits_ReleasesApplicationLease(bool retainOtherConnection)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "failed-cleanup.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var identity = Assert.IsType<ExcelProcessIdentity>(owned.ProcessIdentity);
            using var otherLease = retainOtherConnection ? new ExcelApplicationLease(identity) : null;
            using var process = System.Diagnostics.Process.GetProcessById(identity.ProcessId);
            Assert.False(await CanOtherHostAcquireReviewLeaseAsync(identity));
            owned.BeforeApplicationLeaseReleaseHookForTests = () =>
                throw new InvalidOperationException("Test-only failure after Excel COM cleanup.");
            var close = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.close",
                SessionId = ownedId,
                Args = """{"save":false}"""
            });
            Assert.False(close.Success);
            Assert.Contains("Test-only failure after Excel COM cleanup.",
                close.ErrorMessage + close.InnerError, StringComparison.Ordinal);
            await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(15));
            Assert.True(OwnedProcessGuard.TryConfirmExited(identity));
            Assert.True(SpinWait.SpinUntil(
                () => ((IExcelBatchTeardownState)owned).TryConfirmOwnedProcessTeardown(),
                TimeSpan.FromSeconds(5)));
            await CloseSessionAsync(service, ownedId, save: false);
            sessions.TryRemove(ownedId, out _);
            Assert.Equal(!retainOtherConnection, await CanOtherHostAcquireReviewLeaseAsync(identity));
            otherLease?.Dispose();
            Assert.True(await CanOtherHostAcquireReviewLeaseAsync(identity));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CloseOwnedWorkbook_RestoresAlertsForUnsavedNeighbor(bool retireClosedWorkbook)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "alerts-owner.xlsx");
            var ownedId = await CreateSessionAsync(service, path);
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var name = owned.Execute((context, _) =>
            {
                Assert.False(context.App.DisplayAlerts);
                return AddReviewWorkbook(context.App, "keep unsaved neighbor");
            });
            if (retireClosedWorkbook)
            {
                owned.Execute((context, _) =>
                {
                    context.Book.Close(SaveChanges: false);
                    Excel.Workbooks? workbooks = null;
                    Excel.Workbook? neighborBook = null;
                    try
                    {
                        workbooks = context.App.Workbooks;
                        neighborBook = workbooks[name];
                        neighborBook.SaveAs(path);
                    }
                    finally
                    {
                        ComUtilities.Release(ref neighborBook);
                        ComUtilities.Release(ref workbooks);
                    }
                });
                name = "alerts-owner.xlsx";
            }
            var neighborId = await AttachReviewWorkbookAsync(service,
                await DiscoverReviewWorkbookAsync(service, name));
            sessions[neighborId] = 0;
            var neighbor = Assert.IsType<AttachedWorkbookBatch>(service.SessionManager.GetSession(neighborId));
            await WriteMarkerAsync(service, neighborId, "keep unsaved neighbor");
            if (!retireClosedWorkbook)
            {
                await CloseSessionAsync(service, ownedId, save: false);
            }
            sessions.TryRemove(ownedId, out _);
            neighbor.Execute((context, _) =>
            {
                Assert.True(context.App.DisplayAlerts);
                Assert.False(context.Book.Saved);
            });
            Assert.Equal("keep unsaved neighbor", await ReadMarkerAsync(service, neighborId));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Shutdown_AttachedApplication_PreservesExistingAlerts(bool displayAlerts)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "alerts-unchanged.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            owned.Execute((context, _) => context.App.DisplayAlerts = displayAlerts);
            using var observer = new ExcelMcpService();
            var observedId = await AttachReviewWorkbookAsync(observer,
                await DiscoverReviewWorkbookAsync(observer, "alerts-unchanged.xlsx"));
            await WriteMarkerAsync(observer, observedId, "keep original alerts");
            observer.Dispose();
            owned.Execute((context, _) => Assert.Equal(displayAlerts, context.App.DisplayAlerts));
            Assert.Equal("keep original alerts", await ReadMarkerAsync(service, ownedId));
        });
    }

    [Fact]
    public async Task Close_QuitRejectedThenWorkbookOpened_PreservesLateUnsavedWorkbook()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "quit-backoff.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            Excel.Application? application = null;
            owned.Execute((context, _) => application = context.App);
            var attempts = 0;
            string? lateName = null;
            ExcelShutdownService.BeforeQuitHookForTests = () =>
            {
                if (++attempts == 1)
                {
#pragma warning disable CA2201 // Exercise the real Quit retry policy with a transient COM rejection.
                    throw new System.Runtime.InteropServices.COMException(
                        "Test-only rejected Quit.", ResiliencePipelines.RPC_E_CALL_REJECTED);
#pragma warning restore CA2201
                }
                lateName = AddReviewWorkbook(application!, "opened during quit retry");
            };
            try
            {
                await CloseSessionAsync(service, ownedId, save: false);
                sessions.TryRemove(ownedId, out _);
            }
            finally
            {
                ExcelShutdownService.BeforeQuitHookForTests = null;
                application = null;
            }
            Assert.Equal(2, attempts);
            var survivor = await DiscoverReviewWorkbookAsync(service, Assert.IsType<string>(lateName));
            Assert.False(survivor.GetProperty("saved").GetBoolean());
            var survivorId = await AttachReviewWorkbookAsync(service, survivor);
            sessions[survivorId] = 0;
            Assert.Equal("opened during quit retry", await ReadMarkerAsync(service, survivorId));
            service.SessionManager.GetSession(survivorId)!.Execute((context, _) =>
                Assert.True(context.App.DisplayAlerts));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Discovery_OlderSnapshot_DoesNotRetireNewlyAttachedWorkbook(bool ownedClose)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "snapshot-owner.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var firstName = owned.Execute((context, _) => AddReviewWorkbook(context.App, "first neighbor"));
            var firstId = await AttachReviewWorkbookAsync(service,
                await DiscoverReviewWorkbookAsync(service, firstName));
            sessions[firstId] = 0;
            var application = Assert.IsType<AttachedWorkbookBatch>(
                service.SessionManager.GetSession(firstId)).Owner;
            using var enumerated = new ManualResetEventSlim();
            using var release = new ManualResetEventSlim();
            var scans = 0;
            Action<ExcelBatch> pause = scannedApplication =>
            {
                if (!ReferenceEquals(application, scannedApplication)
                    || Interlocked.Increment(ref scans) != 1)
                    return;
                enumerated.Set();
                Assert.True(release.Wait(TimeSpan.FromSeconds(30)));
            };
            if (ownedClose)
                service.SessionManager.AfterOwnedCloseDiscoveryHookForTests = pause;
            else
                service.SessionManager.AfterApplicationDiscoveryHookForTests = pause;
            var older = Task.Run(async () => ownedClose
                ? await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.close",
                    SessionId = ownedId,
                    Args = """{"save":false}"""
                })
                : await service.ProcessAsync(new ServiceRequest { Command = "session.discover" }));
            string? newId = null;
            JsonElement newSelection = default;
            try
            {
                Assert.True(enumerated.Wait(TimeSpan.FromSeconds(10)));
                var newName = owned.Execute((context, _) => AddReviewWorkbook(context.App, "new unsaved neighbor"));
                newSelection = await DiscoverReviewWorkbookAsync(service, newName);
                newId = await AttachReviewWorkbookAsync(service, newSelection);
                sessions[newId] = 0;
            }
            finally
            {
                release.Set();
                service.SessionManager.AfterApplicationDiscoveryHookForTests = null;
                service.SessionManager.AfterOwnedCloseDiscoveryHookForTests = null;
                RequireSuccess(await older);
                if (ownedClose)
                    sessions.TryRemove(ownedId, out _);
            }
            var session = Assert.IsType<AttachedWorkbookBatch>(
                service.SessionManager.GetSession(Assert.IsType<string>(newId)));
            Assert.Equal("new unsaved neighbor", await ReadMarkerAsync(service, newId));
            session.Execute((context, _) => Assert.False(context.Book.Saved));
            Assert.Equal(newId, await AttachReviewWorkbookAsync(service, newSelection));
            Assert.Equal(newId, await AttachReviewWorkbookAsync(service,
                await DiscoverReviewWorkbookAsync(service, session.WorkbookName)));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Attach_AfterExcelSaveAsWithoutRediscovery_UpdatesExistingSessionPath(bool initiallySaved)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "reattach-owner.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var name = owned.Execute((context, _) => AddReviewWorkbook(context.App));
            var selection = await DiscoverReviewWorkbookAsync(service, name);
            var attachedId = await AttachReviewWorkbookAsync(service, selection);
            sessions[attachedId] = 0;
            var attached = Assert.IsType<AttachedWorkbookBatch>(service.SessionManager.GetSession(attachedId));
            var oldPath = Path.Join(directory, "attached-before.xlsx");
            if (initiallySaved)
            {
                attached.Execute((context, _) => context.Book.SaveAs(oldPath));
                selection = await DiscoverReviewWorkbookAsync(service, "attached-before.xlsx");
            }
            var newPath = Path.Join(directory, "attached-after.xlsx");
            attached.Execute((context, _) => context.Book.SaveAs(newPath));
            await WriteMarkerAsync(service, attachedId, "unsaved after rename");

            Assert.Equal(attachedId, await AttachReviewWorkbookAsync(service, selection));
            Assert.True(service.SessionManager.TryGetFilePath(attachedId, out var trackedPath));
            Assert.Equal(newPath, trackedPath, ignoreCase: true);
            Assert.Equal(newPath, attached.WorkbookPath, ignoreCase: true);
            Assert.Equal("unsaved after rename", await ReadMarkerAsync(service, attachedId));
            Assert.Equal(newPath, Assert.Single(service.SessionManager.GetActiveSessions(),
                session => session.SessionId == attachedId).FilePath, ignoreCase: true);
            var reservation = service.SessionManager.ReserveSessionFilePath(ownedId,
                initiallySaved ? oldPath : Path.Join(directory, "unclaimed.xlsx"));
            service.SessionManager.ReleaseSessionFilePathReservation(ownedId, reservation);
            Assert.Throws<InvalidOperationException>(() =>
                service.SessionManager.ReserveSessionFilePath(ownedId, newPath));
        });
    }

    [Theory]
    [InlineData(false, true)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    public async Task Shutdown_RunningAttachedCallback_StopsWithoutDisturbingCallback(
        bool tracked, bool retainSession)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "pending-owner.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var name = owned.Execute((context, _) => AddReviewWorkbook(context.App, "pending unsaved"));
            var selection = await DiscoverReviewWorkbookAsync(service, name);
            var attachedId = await AttachReviewWorkbookAsync(service, selection);
            sessions[attachedId] = 0;
            var attached = Assert.IsType<AttachedWorkbookBatch>(service.SessionManager.GetSession(attachedId));
            await CloseSessionAsync(service, ownedId, save: false);
            sessions.TryRemove(ownedId, out _);
            // Closing the owned session leaves Excel visible for the remaining workbook; hide it so the
            // test proves that shutdown itself reveals the copy it leaves open.
            attached.Owner.Execute((context, _) => context.App.Visible = false);
            var identity = Assert.IsType<ExcelProcessIdentity>(attached.Owner.ProcessIdentity);
            if (!tracked)
                SessionManager.UntrackExcelProcess(identity);
            if (!retainSession)
            {
                await CloseSessionAsync(service, attachedId, save: false);
                sessions.TryRemove(attachedId, out _);
            }
            using var entered = new ManualResetEventSlim();
            using var release = new ManualResetEventSlim();
            using var createPaused = new ManualResetEventSlim();
            using var resumeCreate = new ManualResetEventSlim();
            service.SessionManager.BeforeNewFileSessionPublishHookForTests = () =>
            {
                createPaused.Set();
                Assert.True(resumeCreate.Wait(TimeSpan.FromSeconds(60)));
            };
            var create = Task.Run(() => service.SessionManager.CreateSessionForNewFile(
                Path.Join(directory, "paused-create.xlsx")));
            var operation = Task.Run(() => attached.Owner.Execute((_, _) =>
            {
                entered.Set();
                Assert.True(release.Wait(TimeSpan.FromSeconds(60)));
                if (retainSession)
                {
                    var context = attached.Owner.GetAttachedWorkbookContext(
                        selection.GetProperty("workbookId").GetString()!);
                    Assert.False(context.Book.Saved);
                }
            }));
            try
            {
                Assert.True(entered.Wait(TimeSpan.FromSeconds(10)));
                Assert.True(createPaused.Wait(TimeSpan.FromSeconds(10)));
                var before = attached.Owner.OperationStatus;
                Assert.True(before.ExecutionPending);
                Assert.False(IsExcelWindowVisible(identity, attached.Owner.Logger));
                var report = service.RequestShutdown();
                Assert.NotNull(report);
                Assert.Contains("its workbooks were not closed: Cannot disconnect: Excel work is still pending",
                    report, StringComparison.Ordinal);
                var sessionCount = service.SessionManager.ActiveSessionCount;
                resumeCreate.Set();
                var createFailure = await Assert.ThrowsAsync<InvalidOperationException>(() => create);
                Assert.IsType<ObjectDisposedException>(createFailure.InnerException);
                Assert.Equal(sessionCount, service.SessionManager.ActiveSessionCount);
                Assert.True(SpinWait.SpinUntil(() => IsExcelWindowVisible(identity, attached.Owner.Logger),
                    TimeSpan.FromSeconds(5)), "Shutdown must reveal the Excel copy it leaves as it is.");
                if (retainSession)
                    Assert.Same(attached, service.SessionManager.GetSession(attachedId));
                Assert.True(service.SessionManager.HasActiveConnections);
                Assert.True(attached.Owner.OperationStatus.ExecutionPending);
                Assert.Equal(before.State, attached.Owner.OperationStatus.State);
                Assert.Equal("SessionUnavailable",
                    (await service.ProcessAsync(new ServiceRequest { Command = "session.list" })).ErrorCategory);
            }
            finally
            {
                release.Set();
                resumeCreate.Set();
                service.SessionManager.BeforeNewFileSessionPublishHookForTests = null;
                try
                {
                    await operation;
                }
                finally
                {
                    if (!tracked)
                        SessionManager.TrackExcelProcess(identity);
                }
            }
            if (retainSession)
            {
                Assert.Equal("pending unsaved", ReadMarker(attached));
            }
            else
            {
                Assert.Equal(0, service.SessionManager.ActiveSessionCount);
                var discovery = attached.Owner.DiscoverAttachedWorkbooks();
                Assert.Empty(discovery.Errors);
                Assert.Empty(discovery.Workbooks);
            }
        });
    }

    private static bool IsExcelWindowVisible(ExcelProcessIdentity identity, Microsoft.Extensions.Logging.ILogger logger) =>
        ExcelDialogProbe.Inspect(identity, logger).Windows.Any(window =>
            window.ProcessId == identity.ProcessId && window.ClassName == "XLMAIN" && window.Visible);

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task DiscoveryConnection_AfterExactProcessExit_ReleasesConnectionAndSessions(bool rediscover)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "exited-owner.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var name = owned.Execute((context, _) => AddReviewWorkbook(context.App));
            var selection = await DiscoverReviewWorkbookAsync(service, name);
            var attachedId = await AttachReviewWorkbookAsync(service, selection);
            sessions[attachedId] = 0;
            var attached = Assert.IsType<AttachedWorkbookBatch>(service.SessionManager.GetSession(attachedId));
            var connection = attached.Owner;
            var identity = Assert.IsType<ExcelProcessIdentity>(connection.ProcessIdentity);
            await CloseSessionAsync(service, ownedId, save: false);
            sessions.TryRemove(ownedId, out _);
            using var process = System.Diagnostics.Process.GetProcessById(identity.ProcessId);
            attached.Execute((context, _) =>
            {
                context.Book.Close(SaveChanges: false);
                context.App.Quit();
            });
            // Model a detach completing after its caller stopped waiting: the manager still
            // retains the connection and selections when the now-empty application exits.
            connection.DetachApplication();
            await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(15));
            Assert.True(OwnedProcessGuard.TryConfirmExited(identity));

            if (rediscover)
            {
                var result = await service.SessionManager.DiscoverOpenWorkbooksAsync();
                Assert.DoesNotContain(result.Errors, error => error.ProcessId == identity.ProcessId);
                Assert.DoesNotContain(result.Workbooks, workbook => workbook.ProcessId == identity.ProcessId);
                Assert.False(service.SessionManager.TryGetFilePath(attachedId, out _));
                Assert.Throws<ArgumentException>(() => service.SessionManager.AttachWorkbook(
                    selection.GetProperty("workbookId").GetString()!));
            }
            else
            {
                service.Dispose();
            }
            Assert.False(connection.OperationStatus.ExecutionPending);
            Assert.Equal(0, service.SessionManager.ActiveSessionCount);
            Assert.False(service.SessionManager.HasActiveConnections);
            Assert.True(await CanOtherHostAcquireReviewLeaseAsync(identity));
            sessions.TryRemove(attachedId, out _);
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Shutdown_CompletedAttachedCallbackAfterExpiredWait_DetachesWithoutSaving(bool closeWorkbook)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "expired-wait-owner.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var name = owned.Execute((context, _) => AddReviewWorkbook(context.App, "preserve late work"));
            var selection = await DiscoverReviewWorkbookAsync(service, name);
            var attachedId = await AttachReviewWorkbookAsync(service, selection);
            sessions[attachedId] = 0;
            var attached = Assert.IsType<AttachedWorkbookBatch>(service.SessionManager.GetSession(attachedId));
            var identity = Assert.IsType<ExcelProcessIdentity>(attached.Owner.ProcessIdentity);
            await CloseSessionAsync(service, ownedId, save: false);
            sessions.TryRemove(ownedId, out _);
            using var entered = new ManualResetEventSlim();
            using var release = new ManualResetEventSlim();
            using var wait = new CancellationTokenSource();
            var operation = Task.Run(() => attached.Execute((context, _) =>
            {
                entered.Set();
                Assert.True(release.Wait(TimeSpan.FromSeconds(15)));
                if (closeWorkbook)
                    context.Book.Close(SaveChanges: false);
            }, wait.Token));
            try
            {
                Assert.True(entered.Wait(TimeSpan.FromSeconds(10)));
                wait.Cancel();
                await Assert.ThrowsAnyAsync<OperationCanceledException>(() => operation);
                Assert.True(attached.Owner.OperationStatus.ExecutionPending);
            }
            finally
            {
                release.Set();
            }
            Assert.True(SpinWait.SpinUntil(
                () => !attached.Owner.OperationStatus.ExecutionPending, TimeSpan.FromSeconds(10)));
            Assert.True(attached.HasTimedOutOperation);
            service.SessionManager.Dispose();
            sessions.TryRemove(attachedId, out _);
            Assert.False(service.SessionManager.HasActiveConnections);
            if (closeWorkbook)
            {
                Assert.True(OwnedProcessGuard.TryConfirmExited(identity),
                    $"The empty test-owned Excel process did not exit: {identity}.");
                return;
            }

            using var verifier = new ExcelMcpService();
            var survivor = await DiscoverReviewWorkbookAsync(verifier, name);
            Assert.False(survivor.GetProperty("saved").GetBoolean());
            var survivorId = await AttachReviewWorkbookAsync(verifier, survivor);
            try
            {
                Assert.Equal("preserve late work", await ReadMarkerAsync(verifier, survivorId));
            }
            finally
            {
                await CloseSessionAsync(verifier, survivorId, save: false);
            }
        });
    }

    [Fact]
    public async Task OwnedApplication_HoldsLeaseAgainstOtherHostsUntilItsLastConnectionDetaches()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var sessionId = await CreateSessionAsync(service, Path.Join(directory, "lease-owner.xlsx"));
            sessions[sessionId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(sessionId));
            var identity = Assert.IsType<ExcelProcessIdentity>(owned.ProcessIdentity);
            Assert.False(await CanOtherHostAcquireReviewLeaseAsync(identity));
            var neighborName = owned.Execute((context, _) => AddReviewWorkbook(context.App));
            var attached = await AttachReviewWorkbookAsync(service,
                await DiscoverReviewWorkbookAsync(service, neighborName));
            sessions[attached] = 0;
            await CloseSessionAsync(service, sessionId, save: false);
            sessions.TryRemove(sessionId, out _);
            Assert.False(await CanOtherHostAcquireReviewLeaseAsync(identity));
            await CloseSessionAsync(service, attached, save: false);
            sessions.TryRemove(attached, out _);
            service.SessionManager.Dispose();
            Assert.True(await CanOtherHostAcquireReviewLeaseAsync(identity));
        });
    }

    private static async Task<bool> CanOtherHostAcquireReviewLeaseAsync(ExcelProcessIdentity identity)
    {
        var name = $@"Local\ExcelMcp.Application.{identity.ProcessId}.{identity.StartedAtUtcFileTime}";
        var start = new System.Diagnostics.ProcessStartInfo("pwsh")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        start.ArgumentList.Add("-NoProfile");
        start.ArgumentList.Add("-Command");
        start.ArgumentList.Add(
            $"$mutex = [Threading.Mutex]::new($false, '{name}'); " +
            "try { $acquired = $mutex.WaitOne(0); " +
            "if ($acquired) { $mutex.ReleaseMutex() }; [Console]::Write($acquired) } " +
            "finally { $mutex.Dispose() }");
        using var process = System.Diagnostics.Process.Start(start)!;
        var output = process.StandardOutput.ReadToEndAsync();
        var errors = process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(15));
        Assert.Equal(0, process.ExitCode);
        Assert.Equal(string.Empty, await errors);
        return bool.Parse(await output);
    }

    [Fact]
    public async Task Attach_StaleOwnedSession_DoesNotHoldPathLockWhileResolvingSession()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "stale-lock.xlsx");
            var original = await CreateSessionAsync(service, path);
            sessions[original] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(original));
            var neighborName = owned.Execute((context, _) => AddReviewWorkbook(context.App));
            var neighbor = await DiscoverReviewWorkbookAsync(service, neighborName);
            owned.Execute((context, _) =>
            {
                Excel.Workbooks? workbooks = null;
                Excel.Workbook? replacement = null;
                try
                {
                    workbooks = context.App.Workbooks;
                    replacement = workbooks[neighborName];
                    context.Book.Close(SaveChanges: false);
                    replacement.SaveAs(path);
                }
                finally
                {
                    ComUtilities.Release(ref replacement);
                    ComUtilities.Release(ref workbooks);
                }
            });
            Task? operation = null;
            service.SessionManager.BeforeStaleOwnedSessionLockHookForTests = () =>
            {
                operation = Task.Run(() =>
                {
                    Assert.True(service.SessionManager.TryBeginOperation(original, out _, out var error), error);
                    service.SessionManager.EndOperation(original);
                });
                operation.WaitAsync(TimeSpan.FromSeconds(3)).GetAwaiter().GetResult();
            };
            try
            {
                var attached = await AttachReviewWorkbookAsync(service, neighbor);
                sessions.TryRemove(original, out _);
                sessions[attached] = 0;
                Assert.NotEqual(original, attached);
                Assert.Null(service.SessionManager.GetSession(original));
                Assert.Equal(path, service.SessionManager.GetSession(attached)!.WorkbookPath, ignoreCase: true);
                await WriteMarkerAsync(service, attached, "replacement remains usable");
                Assert.Equal("replacement remains usable", await ReadMarkerAsync(service, attached));
            }
            finally
            {
                service.SessionManager.BeforeStaleOwnedSessionLockHookForTests = null;
                if (operation is not null)
                    await operation;
                if (service.SessionManager.GetSession(original) is not null)
                {
                    var replacement = await AttachReviewWorkbookAsync(service,
                        await DiscoverReviewWorkbookAsync(service, "stale-lock.xlsx"));
                    sessions.TryRemove(original, out _);
                    sessions[replacement] = 0;
                }
            }
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Discover_OwnedWorkbookRenamedInExcel_ReusesOriginalSessionAndReleasesOldPath(bool reserveNewPath)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var oldPath = Path.Join(directory, "owned-before.xlsx");
            var newPath = Path.Join(directory, "owned-after.xlsx");
            var sessionId = await CreateSessionAsync(service, oldPath);
            sessions[sessionId] = 0;
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(sessionId));
            var before = await DiscoverReviewWorkbookAsync(service, "owned-before.xlsx");
            string? blocker = null;
            if (reserveNewPath)
            {
                blocker = await CreateSessionAsync(service, Path.Join(directory, "reservation-owner.xlsx"));
                sessions[blocker] = 0;
                service.SessionManager.ReserveSessionFilePath(blocker, newPath);
            }
            batch.Execute((context, _) => context.Book.SaveAs(newPath));
            await WriteMarkerAsync(service, sessionId, "renamed unsaved");

            if (blocker is not null)
            {
                try
                {
                    var incomplete = await service.ProcessAsync(new ServiceRequest { Command = "session.discover" });
                    RequireSuccess(incomplete);
                    using var result = JsonDocument.Parse(incomplete.Result!);
                    Assert.False(result.RootElement.GetProperty("complete").GetBoolean());
                    Assert.Contains(result.RootElement.GetProperty("errors").EnumerateArray(),
                        item => item.GetProperty("message").GetString()!.Contains("reserved", StringComparison.Ordinal));
                    Assert.Equal(oldPath, batch.WorkbookPath, ignoreCase: true);
                    Assert.Equal("renamed unsaved", await ReadMarkerAsync(service, sessionId));
                }
                finally
                {
                    service.SessionManager.ReleaseSessionFilePathReservation(blocker, newPath);
                }
                await CloseSessionAsync(service, blocker, save: false);
                sessions.TryRemove(blocker, out _);
            }

            var after = await DiscoverReviewWorkbookAsync(service, "owned-after.xlsx");
            Assert.Equal(before.GetProperty("workbookId").GetString(), after.GetProperty("workbookId").GetString());
            Assert.Equal(sessionId, after.GetProperty("sessionId").GetString());
            Assert.False(after.GetProperty("saved").GetBoolean());
            var attached = await AttachReviewWorkbookAsync(service, after);
            Assert.Equal(sessionId, attached);
            Assert.Single(service.SessionManager.GetActiveSessions());
            Assert.True(service.SessionManager.TryGetFilePath(sessionId, out var path));
            Assert.Equal(newPath, path, ignoreCase: true);
            Assert.Equal(newPath, batch.WorkbookPath, ignoreCase: true);
            Assert.Equal("renamed unsaved", await ReadMarkerAsync(service, sessionId));

            var reopened = await OpenSessionAsync(service, oldPath);
            sessions[reopened] = 0;
            Assert.NotEqual(sessionId, reopened);
            await CloseSessionAsync(service, reopened, save: false);
            sessions.TryRemove(reopened, out _);
        });
    }

    [Fact]
    public async Task Shutdown_WorkbookOpenedAfterEmptyCheck_PreservesNewUnsavedWorkbook()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var sessionId = await CreateSessionAsync(service, Path.Join(directory, "empty-owner.xlsx"));
            sessions[sessionId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(sessionId));
            var neighborName = owned.Execute((context, _) => AddReviewWorkbook(context.App));
            var neighbor = await AttachReviewWorkbookAsync(service,
                await DiscoverReviewWorkbookAsync(service, neighborName));
            sessions[neighbor] = 0;
            var connection = Assert.IsType<AttachedWorkbookBatch>(
                service.SessionManager.GetSession(neighbor)).Owner;
            await CloseSessionAsync(service, sessionId, save: false);
            sessions.TryRemove(sessionId, out _);
            await CloseSessionAsync(service, neighbor, save: false);
            sessions.TryRemove(neighbor, out _);

            string? newName = null;
            connection.BeforeEmptyApplicationDetachHookForTests = app =>
                newName = AddReviewWorkbook((Excel.Application)app, "late unsaved");
            try
            {
                service.SessionManager.Dispose();
            }
            finally
            {
                connection.BeforeEmptyApplicationDetachHookForTests = null;
            }

            using var verifier = new ExcelMcpService();
            var survivor = await DiscoverReviewWorkbookAsync(verifier, Assert.IsType<string>(newName));
            Assert.False(survivor.GetProperty("saved").GetBoolean());
            var survivorSession = await AttachReviewWorkbookAsync(verifier, survivor);
            try
            {
                Assert.Equal("late unsaved", await ReadMarkerAsync(verifier, survivorSession));
            }
            finally
            {
                await CloseSessionAsync(verifier, survivorSession, save: false);
            }
        });
    }

    [Fact]
    public async Task Execute_OwnedAndAttachedCallbacks_SerializeStateRestoration()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var sessionId = await CreateSessionAsync(service, Path.Join(directory, "serialized-owner.xlsx"));
            sessions[sessionId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(sessionId));
            var neighborName = owned.Execute((context, _) => AddReviewWorkbook(context.App));
            var attachedId = await AttachReviewWorkbookAsync(service,
                await DiscoverReviewWorkbookAsync(service, neighborName));
            sessions[attachedId] = 0;
            var attached = Assert.IsType<AttachedWorkbookBatch>(service.SessionManager.GetSession(attachedId));
            using var entered = new ManualResetEventSlim();
            using var release = new ManualResetEventSlim();
            using var queued = new ManualResetEventSlim();
            using var attachedEntered = new ManualResetEventSlim();
            Task? first = null;
            Task? second = null;
            try
            {
                // Exercise the complete callback boundary, including application-state restoration.
                first = Task.Run(() => owned.Execute((context, _) =>
                {
                    context.App.Calculation = Excel.XlCalculation.xlCalculationManual;
                    entered.Set();
                    try
                    {
                        Assert.True(release.Wait(TimeSpan.FromSeconds(15)));
                    }
                    finally
                    {
                        context.App.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                    }
                }));
                Assert.True(entered.Wait(TimeSpan.FromSeconds(10)));
                ExcelBatch.WorkItemQueuedHookForTests = queued.Set;
                second = Task.Run(() => attached.Execute((context, _) =>
                {
                    attachedEntered.Set();
                    context.App.Calculation = Excel.XlCalculation.xlCalculationManual;
                }));
                Assert.True(queued.Wait(TimeSpan.FromSeconds(10)));
                Assert.False(attachedEntered.Wait(TimeSpan.FromMilliseconds(500)),
                    "The attached callback ran before the owned callback restored application state.");
            }
            finally
            {
                release.Set();
                ExcelBatch.WorkItemQueuedHookForTests = null;
                if (first is not null)
                    await first;
                if (second is not null)
                    await second;
            }
            Assert.True(attachedEntered.IsSet);
            attached.Execute((context, _) =>
            {
                Assert.Equal(Excel.XlCalculation.xlCalculationManual, context.App.Calculation);
                context.App.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            });
        });
    }

    [Fact]
    public async Task Trace_AttachedUnownedProcess_UsesConnectionIdentityWithoutClaimingOwnership()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var sessionId = await CreateSessionAsync(service, Path.Join(directory, "trace-owner.xlsx"));
            sessions[sessionId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(sessionId));
            var neighborName = owned.Execute((context, _) => AddReviewWorkbook(context.App));
            var attached = await AttachReviewWorkbookAsync(service,
                await DiscoverReviewWorkbookAsync(service, neighborName));
            sessions[attached] = 0;
            var identity = Assert.IsType<ExcelProcessIdentity>(owned.ProcessIdentity);
            SessionManager.UntrackExcelProcess(identity);
            try
            {
                foreach (var command in ReviewTraceCommands)
                {
                    var response = await service.ProcessAsync(new ServiceRequest
                    {
                        Command = command,
                        Args = JsonSerializer.Serialize(new { workbookSessionId = attached }, ServiceProtocol.JsonOptions)
                    });
                    RequireSuccess(response);
                    using var result = JsonDocument.Parse(response.Result!);
                    Assert.True(result.RootElement.GetProperty("success").GetBoolean());
                    Assert.Equal(identity.ProcessId, result.RootElement.GetProperty("processId").GetInt32());
                    Assert.DoesNotContain(SessionManager.GetTrackedExcelProcesses(), item => item == identity);
                }
            }
            finally
            {
                SessionManager.TrackExcelProcess(identity);
            }
        });
    }

    [Fact]
    public async Task VbaRun_AttachedWorkbookWithDuplicateProcedure_TargetsSelectedWorkbook()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var created = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.create",
                Args = JsonSerializer.Serialize(new
                {
                    filePath = Path.Join(directory, "active.xlsm")
                }, ServiceProtocol.JsonOptions)
            });
            var activeId = GetSessionId(created);
            sessions[activeId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(activeId));
            var selectedName = owned.Execute((context, _) =>
            {
                Excel.Workbooks? workbooks = null;
                Excel.Workbook? workbook = null;
                try
                {
                    workbooks = context.App.Workbooks;
                    workbook = workbooks.Add();
                    workbook.SaveAs(Path.Join(directory, "selected's.xlsm"),
                        ComInteropConstants.XlOpenXmlWorkbookMacroEnabled);
                    return workbook.Name;
                }
                finally
                {
                    ComUtilities.Release(ref workbook);
                    ComUtilities.Release(ref workbooks);
                }
            });
            var selectedId = await AttachReviewWorkbookAsync(service,
                await DiscoverReviewWorkbookAsync(service, selectedName));
            sessions[selectedId] = 0;
            var otherName = owned.Execute((context, _) =>
            {
                Excel.Workbooks? workbooks = null;
                Excel.Workbook? workbook = null;
                try
                {
                    workbooks = context.App.Workbooks;
                    workbook = workbooks.Add();
                    workbook.SaveAs(Path.Join(directory, "other.xlsm"),
                        ComInteropConstants.XlOpenXmlWorkbookMacroEnabled);
                    return workbook.Name;
                }
                finally
                {
                    ComUtilities.Release(ref workbook);
                    ComUtilities.Release(ref workbooks);
                }
            });
            var otherId = await AttachReviewWorkbookAsync(service,
                await DiscoverReviewWorkbookAsync(service, otherName));
            sessions[otherId] = 0;
            foreach (var id in new[] { selectedId, otherId, activeId })
            {
                RequireSuccess(await service.ProcessAsync(new ServiceRequest
                {
                    Command = "vba.import",
                    SessionId = id,
                    Args = JsonSerializer.Serialize(new
                    {
                        moduleName = "Module1",
                        vbaCode = "Sub Stamp()\nThisWorkbook.Worksheets(1).Range(\"A1\").Value2 = \"macro ran\"\nEnd Sub"
                    }, ServiceProtocol.JsonOptions)
                }));
                await WriteMarkerAsync(service, id, id == selectedId ? "selected baseline" : "active baseline");
            }
            owned.Execute((context, _) => context.Book.Activate());
            RequireSuccess(await service.ProcessAsync(new ServiceRequest
            {
                Command = "vba.run",
                SessionId = selectedId,
                Args = """{"procedureName":"Module1.Stamp"}"""
            }));
            Assert.Equal("macro ran", await ReadMarkerAsync(service, selectedId));
            Assert.Equal("active baseline", await ReadMarkerAsync(service, activeId));
            Assert.Equal("active baseline", await ReadMarkerAsync(service, otherId));
            RequireSuccess(await service.ProcessAsync(new ServiceRequest
            {
                Command = "vba.run",
                SessionId = otherId,
                Args = """{"procedureName":"Module1.Stamp"}"""
            }));
            Assert.Equal("macro ran", await ReadMarkerAsync(service, otherId));
            Assert.Equal("active baseline", await ReadMarkerAsync(service, activeId));
            var foreignProcedure = await service.ProcessAsync(new ServiceRequest
            {
                Command = "vba.run",
                SessionId = selectedId,
                Args = """{"procedureName":"active.xlsm!Module1.Stamp"}"""
            });
            Assert.False(foreignProcedure.Success);
            Assert.Contains("unqualified procedure name", foreignProcedure.ErrorMessage, StringComparison.Ordinal);
            Assert.Equal("active baseline", await ReadMarkerAsync(service, activeId));
        });
    }

    private static string AddReviewWorkbook(Excel.Application app, string? marker = null)
    {
        Excel.Workbooks? workbooks = null;
        Excel.Workbook? workbook = null;
        Excel.Sheets? sheets = null;
        Excel.Worksheet? sheet = null;
        Excel.Range? cell = null;
        try
        {
            workbooks = app.Workbooks;
            workbook = workbooks.Add();
            if (marker is not null)
            {
                sheets = workbook.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                cell = sheet.Range["A1"];
                cell.Value2 = marker;
            }
            return workbook.Name;
        }
        finally
        {
            ComUtilities.Release(ref cell);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref sheets);
            ComUtilities.Release(ref workbook);
            ComUtilities.Release(ref workbooks);
        }
    }

    private static async Task<JsonElement> DiscoverReviewWorkbookAsync(ExcelMcpService service, string name)
    {
        var response = await service.ProcessAsync(new ServiceRequest { Command = "session.discover" });
        RequireSuccess(response);
        using var result = JsonDocument.Parse(response.Result!);
        return Assert.Single(result.RootElement.GetProperty("workbooks").EnumerateArray(),
            item => item.GetProperty("name").GetString() == name).Clone();
    }

    private static async Task<string> AttachReviewWorkbookAsync(ExcelMcpService service, JsonElement workbook)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.attach",
            Args = JsonSerializer.Serialize(new
            {
                workbookId = workbook.GetProperty("workbookId").GetString()
            }, ServiceProtocol.JsonOptions)
        });
        return GetSessionId(response);
    }
}
