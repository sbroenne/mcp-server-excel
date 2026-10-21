using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    private static readonly string[] ReviewTraceCommands = ["powerquerytrace.list", "powerquerytrace.summarize"];

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
    public async Task Shutdown_RunningAttachedCallback_RetainsConnectionUntilCallbackCompletes(
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
                var before = attached.Owner.OperationStatus;
                Assert.True(before.ExecutionPending);
                Assert.Throws<AggregateException>(service.Dispose);
                if (retainSession)
                    Assert.Same(attached, service.SessionManager.GetSession(attachedId));
                Assert.True(service.SessionManager.HasActiveConnections);
                Assert.True(attached.Owner.OperationStatus.ExecutionPending);
                Assert.Equal(before.State, attached.Owner.OperationStatus.State);
            }
            finally
            {
                release.Set();
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
                Assert.Equal("pending unsaved", await ReadMarkerAsync(service, attachedId));
            }
            else
            {
                RequireSuccess(await service.ProcessAsync(new ServiceRequest { Command = "session.list" }));
                Assert.Equal(0, service.SessionManager.ActiveSessionCount);
                var discovery = attached.Owner.DiscoverAttachedWorkbooks();
                Assert.Empty(discovery.Errors);
                Assert.Empty(discovery.Workbooks);
            }
        });
    }

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
