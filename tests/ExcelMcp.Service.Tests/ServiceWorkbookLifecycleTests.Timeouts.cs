using System.Diagnostics;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Workbook;
using Sbroenne.ExcelMcp.Service.Workers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    private static readonly string[] PrivateWorkerTraceReadCommands = ["powerquerytrace.list", "powerquerytrace.summarize"];

    private static readonly string[] StatusCommands = ["service.ping", "service.status", "session.list"];

    [Fact]
    public async Task Status_DoesNotWaitForWorkbookLocationBookkeeping()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var session = await CreateSessionAsync(service, Path.Join(directory, "responsive-status.xlsx"));
            sessions[session] = 0;
            var beforeWrite = DateTime.UtcNow;
            await WriteMarkerAsync(service, session, "Unsaved contents stay open");
            var gate = typeof(SessionManager).GetField("_filePathReservationLock",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
                .GetValue(service.SessionManager)!;
            using var entered = new ManualResetEventSlim();
            using var release = new ManualResetEventSlim();
            var bookkeeping = Task.Run(() =>
            {
                lock (gate)
                {
                    entered.Set();
                    Assert.True(release.Wait(TimeSpan.FromSeconds(30)));
                }
            });
            var pipe = $"excelmcp-status-bookkeeping-{Guid.NewGuid():N}";
            var host = service.RunAsync(pipe);
            var failure = await Record.ExceptionAsync(async () =>
            {
                Assert.True(entered.Wait(TimeSpan.FromSeconds(5)));
                foreach (var command in StatusCommands)
                {
                    using var client = new ServiceClient(pipe);
                    var response = await client.SendAsync(new ServiceRequest { Command = command })
                        .WaitAsync(TimeSpan.FromSeconds(2));
                    RequireSuccess(response);
                    if (command == "session.list")
                    {
                        using var result = JsonDocument.Parse(response.Result!);
                        var entry = Assert.Single(result.RootElement.GetProperty("sessions").EnumerateArray());
                        Assert.Equal(session, entry.GetProperty("sessionId").GetString());
                        var operation = entry.GetProperty("operation");
                        Assert.Equal("range.set-values", operation.GetProperty("command").GetString());
                        Assert.Equal("completed", operation.GetProperty("state").GetString());
                        Assert.False(operation.GetProperty("executionPending").GetBoolean());
                        var started = operation.GetProperty("startedAtUtc").GetDateTime();
                        Assert.Equal(DateTimeKind.Utc, started.Kind);
                        Assert.InRange(started, beforeWrite, DateTime.UtcNow);
                    }
                }
            });
            release.Set();
            await bookkeeping;
            Assert.Equal("Unsaved contents stay open", await ReadMarkerAsync(service, session));
            await CloseSessionAsync(service, session, save: false);
            service.RequestShutdown();
            await host.WaitAsync(TimeSpan.FromSeconds(15));
            sessions.TryRemove(session, out _);
            if (failure is not null)
                ExceptionDispatchInfo.Capture(failure).Throw();
        });
    }

    [Fact]
    public async Task Status_RemainsAvailableWhileExcelReadinessDoesNotAnswer()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var session = await CreateSessionAsync(service, Path.Join(directory, "blocked-excel-status.xlsx"));
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, "Still unsaved");
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(session));
            var startedAt = batch.OperationStatus.StartedAtUtc;
            using var release = new ManualResetEventSlim();
            var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            batch.BeforeRefreshStateReadHookForTests = () =>
            {
                entered.TrySetResult();
                Assert.True(release.Wait(TimeSpan.FromSeconds(30)));
            };
            var pipe = $"excelmcp-blocked-excel-status-{Guid.NewGuid():N}";
            var host = service.RunAsync(pipe);
            var failure = await Record.ExceptionAsync(async () =>
            {
                using var initial = new ServiceClient(pipe);
                RequireSuccess(await initial.SendAsync(new ServiceRequest { Command = "session.list" })
                    .WaitAsync(TimeSpan.FromSeconds(2)));
                await entered.Task.WaitAsync(TimeSpan.FromSeconds(2));
                await Task.WhenAll(StatusCommands.Select(async command =>
                {
                    using var client = new ServiceClient(pipe);
                    var response = await client.SendAsync(new ServiceRequest { Command = command })
                        .WaitAsync(TimeSpan.FromSeconds(2));
                    RequireSuccess(response);
                    if (command != "session.list")
                        return;
                    using var result = JsonDocument.Parse(response.Result!);
                    var entry = Assert.Single(result.RootElement.GetProperty("sessions").EnumerateArray());
                    Assert.Equal(session, entry.GetProperty("sessionId").GetString());
                    Assert.False(entry.GetProperty("canClose").GetBoolean());
                    Assert.Equal("unknown", entry.GetProperty("excelState").GetString());
                    var operation = entry.GetProperty("operation");
                    Assert.Equal("range.set-values", operation.GetProperty("command").GetString());
                    Assert.Equal(startedAt, operation.GetProperty("startedAtUtc").GetDateTime());
                }));
            });
            batch.BeforeRefreshStateReadHookForTests = null;
            release.Set();
            var cleanup = await Record.ExceptionAsync(async () =>
            {
                var timer = Stopwatch.StartNew();
                while (batch.GetRefreshState() != WorkbookRefreshState.Ready)
                {
                    Assert.True(timer.Elapsed < TimeSpan.FromSeconds(10));
                    await Task.Delay(50);
                }
                Assert.Equal("Still unsaved", await ReadMarkerAsync(service, session));
                await CloseSessionAsync(service, session, save: false);
                sessions.TryRemove(session, out _);
                service.RequestShutdown();
                await host.WaitAsync(TimeSpan.FromSeconds(10));
            });
            if (failure is not null && cleanup is not null)
                throw new AggregateException(failure, cleanup);
            if (failure is not null || cleanup is not null)
                ExceptionDispatchInfo.Capture((failure ?? cleanup)!).Throw();
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    [InlineData(true, true)]
    public async Task Test_UncommittedCloseRefusal_AllowsExplicitCloseAfterReadinessReturns(
        bool disposeService, bool saveAs = false)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "validation-readiness.xlsx");
            var initial = await CreateSessionAsync(service, path);
            await WriteMarkerAsync(service, initial, "Validation retained contents");
            await CloseSessionAsync(service, initial, save: true);
            var savedBytes = await File.ReadAllBytesAsync(path);
            ExcelBatch? batch = null;
            string? session = null;
            ExcelBatch.AfterWorkbookOpenHookForTests = (_, _) =>
            {
                session = Assert.Single(service.SessionManager.GetActiveSessions()).SessionId;
                sessions[session] = 0;
                batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(session));
                batch.BeforeRefreshStateReadHookForTests = () =>
                    ExcelBusyException.ThrowIfNotReady(WorkbookRefreshState.Refreshing, "validation-close");
            };
            try
            {
                var response = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.test",
                    Args = JsonSerializer.Serialize(new { filePath = path }, ServiceProtocol.JsonOptions)
                });
                Assert.False(response.Success);
                Assert.Equal("Busy", response.ErrorCategory);
                Assert.Equal(session, response.SessionId);
                Assert.NotNull(batch);
                Assert.False(((IExcelBatchCloseState)batch).ShutdownCommitted);
            }
            finally
            {
                ExcelBatch.AfterWorkbookOpenHookForTests = null;
                if (batch is not null) batch.BeforeRefreshStateReadHookForTests = null;
            }
            Assert.NotNull(session);
            Assert.Equal("Validation retained contents", await ReadMarkerAsync(service, session));
            if (disposeService)
            {
                var identity = Assert.Single(SessionManager.GetTrackedExcelProcesses(),
                    candidate => candidate.ProcessId == batch!.ExcelProcessId);
                Assert.True(batch!.Execute((context, _) => context.Book.ReadOnly));
                var target = Path.Join(directory, "validation-writable.xlsx");
                if (saveAs)
                {
                    RequireSuccess(await service.ProcessAsync(new ServiceRequest
                    {
                        Command = "workbook.save-as",
                        SessionId = session,
                        Args = JsonSerializer.Serialize(new { targetPath = target }, ServiceProtocol.JsonOptions)
                    }));
                    Assert.False(batch.Execute((context, _) => context.Book.ReadOnly));
                    Assert.True(service.SessionManager.TryGetFilePath(session, out var trackedPath));
                    Assert.Equal(target, trackedPath);
                    await WriteMarkerAsync(service, session, "Edits after validation Save As");
                    Assert.False(batch.Execute((context, _) => context.Book.Saved));
                }
                service.Dispose();
                sessions.TryRemove(session, out _);
                Assert.Equal(0, service.SessionManager.ActiveSessionCount);
                Assert.Empty(service.SessionManager.ActiveSessionIds);
                var paths = Assert.IsAssignableFrom<System.Collections.IDictionary>(
                    typeof(SessionManager).GetField("_sessionFilePaths",
                        System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
                        .GetValue(service.SessionManager));
                Assert.Empty(paths);
                Assert.Equal(savedBytes, await File.ReadAllBytesAsync(path));
                if (saveAs)
                {
                    // Unsaved edits after Save As: Excel is left open and visible, nothing saved.
                    Assert.True(OwnedProcessGuard.MatchesLiveProcess(identity));
                    await RunWithCleanupAsync(async (verification, _, reopenedSessions) =>
                    {
                        var reopened = await AttachHandoverAsync(verification, target, identity.ProcessId);
                        reopenedSessions[reopened] = 0;
                        Assert.Equal("Edits after validation Save As", await ReadMarkerAsync(verification, reopened));
                        await CloseSessionAsync(verification, reopened, save: false);
                        reopenedSessions.TryRemove(reopened, out var released);
                    });
                    return;
                }
                Assert.True(OwnedProcessGuard.TryConfirmExited(identity));
                Assert.DoesNotContain(identity, SessionManager.GetTrackedExcelProcesses());
                return;
            }

            await CloseSessionAsync(service, session, save: false);
            sessions.TryRemove(session, out _);
            Assert.False(service.SessionManager.TryGetFilePath(session, out _));
            var reopened = await OpenSessionAsync(service, path);
            sessions[reopened] = 0;
            Assert.Equal("Validation retained contents", await ReadMarkerAsync(service, reopened));
        });
    }

    [Fact]
    public async Task Close_CancelledBeforeExecution_OrdinaryShutdownHandsOverUnsavedChanges()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "cancelled-close.xlsx");
            var session = await CreateSessionAsync(service, path);
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, "Saved baseline");
            RequireSuccess(await service.ProcessAsync(
                new ServiceRequest { Command = "session.save", SessionId = session }));
            await WriteMarkerAsync(service, session, "Latest unsaved changes");
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(session));
            var identity = Assert.Single(SessionManager.GetTrackedExcelProcesses(),
                candidate => candidate.ProcessId == batch.ExcelProcessId);
            using var cancellation = new CancellationTokenSource();
            service.SessionManager.BeforeCloseSessionHookForTests = cancellation.Cancel;
            ServiceResponse response;
            try
            {
                response = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.close",
                    SessionId = session,
                    CallerCancellation = cancellation.Token
                });
            }
            finally
            {
                service.SessionManager.BeforeCloseSessionHookForTests = null;
            }
            Assert.False(response.Success);
            Assert.Equal("Cancelled", response.ErrorCategory);
            Assert.False(((IExcelBatchCloseState)batch).ShutdownCommitted);
            Assert.Same(batch, service.SessionManager.GetSession(session));
            Assert.True(OwnedProcessGuard.MatchesLiveProcess(identity));
            Assert.True(service.SessionManager.TryGetFilePath(session, out var retainedPath));
            Assert.Equal(path, retainedPath);
            Assert.False(batch.Execute((context, _) => (bool)context.Book.Saved));

            service.Dispose();
            sessions.TryRemove(session, out _);
            Assert.True(OwnedProcessGuard.MatchesLiveProcess(identity));
            await RunWithCleanupAsync(async (verification, _, reopenedSessions) =>
            {
                var reopened = await AttachHandoverAsync(verification, path, batch.ExcelProcessId);
                reopenedSessions[reopened] = 0;
                Assert.Equal("Latest unsaved changes", await ReadMarkerAsync(verification, reopened));
                await CloseSessionAsync(verification, reopened, save: false);
                reopenedSessions.TryRemove(reopened, out var removed);
                var disk = await OpenSessionAsync(verification, path);
                reopenedSessions[disk] = 0;
                Assert.Equal("Saved baseline", await ReadMarkerAsync(verification, disk));
            });
        });
    }

    [Fact]
    public async Task PrivateWorker_CancelledRunningSave_RevealsExcelBeforeCompletion()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var registration = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}",
                WorkerProcess.Current, WorkerProcess.Current, Guid.NewGuid().ToString("N"),
                WorkerRegistry.Build, null);
            using var host = new ExcelWorkerHost(registration, service);
            var run = host.RunAsync();
            var worker = new ExcelWorkerClient(registration, WorkerProcess.Current, registration.Generation);
            using var release = new ManualResetEventSlim();
            using var cancellation = new CancellationTokenSource();
            Task<ServiceResponse>? save = null;
            var path = Path.Join(directory, "private-worker-cancelled-save.xlsx");
            var failure = await Record.ExceptionAsync(async () =>
            {
                await worker.InitializeAsync();
                var session = await CreateSessionAsync(service, path);
                sessions[session] = 0;
                await WriteMarkerAsync(service, session, "Saved only by the requested running save");
                var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(session));
                Assert.False(batch.IsExcelVisible);
                using var excel = Process.GetProcessById(batch.ExcelProcessId!.Value);
                Assert.Equal(batch.ProcessIdentity,
                    new ExcelProcessIdentity(excel.Id, excel.StartTime.ToUniversalTime().ToFileTimeUtc()));
                Assert.Equal(IntPtr.Zero, excel.MainWindowHandle);
                var saving = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
                ExcelShutdownService.BeforeSaveHookForTests = () =>
                {
                    saving.TrySetResult();
                    Assert.True(release.Wait(TimeSpan.FromSeconds(30)));
                };
                save = worker.ExecuteAsync(new ServiceRequest
                {
                    Command = "session.save",
                    SessionId = session
                }, cancellation.Token);
                await saving.Task.WaitAsync(TimeSpan.FromSeconds(10));
                foreach (var command in PrivateWorkerTraceReadCommands)
                {
                    var traces = await worker.ExecuteAsync(new ServiceRequest
                    {
                        Command = command,
                        Args = JsonSerializer.Serialize(new { workbookSessionId = session })
                    }, CancellationToken.None).WaitAsync(TimeSpan.FromSeconds(2));
                    RequireSuccess(traces);
                    using var result = JsonDocument.Parse(traces.Result!);
                    Assert.Equal(batch.ExcelProcessId, result.RootElement.GetProperty("processId").GetInt32());
                    Assert.True(batch.OperationStatus.ExecutionPending);
                    Assert.False(save.IsCompleted);
                }
                cancellation.Cancel();
                Assert.True(SpinWait.SpinUntil(() =>
                {
                    excel.Refresh();
                    return batch.OperationStatus.WaitExpired && excel.MainWindowHandle != IntPtr.Zero;
                }, TimeSpan.FromSeconds(3)),
                    $"Wait expired: {batch.OperationStatus.WaitExpired}; window: {batch.OperationStatus.WindowReveal}.");
                Assert.True(batch.OperationStatus.ExecutionPending);
                Assert.False(save.IsCompleted);
                ExcelShutdownService.BeforeSaveHookForTests = null;
                release.Set();
                RequireSuccess(await save.WaitAsync(TimeSpan.FromSeconds(10)));
                Assert.Equal("Saved only by the requested running save", await ReadMarkerAsync(service, session));
                await CloseSessionAsync(service, session, save: false);
                sessions.TryRemove(session, out _);
            });
            ExcelShutdownService.BeforeSaveHookForTests = null;
            release.Set();
            var completion = await Record.ExceptionAsync(async () =>
            {
                if (save is not null)
                    RequireSuccess(await save.WaitAsync(TimeSpan.FromSeconds(10)));
                foreach (var session in sessions.Keys)
                {
                    service.SessionManager.CloseSession(session, save: false);
                    sessions.TryRemove(session, out _);
                }
            });
            var shutdown = await Record.ExceptionAsync(async () =>
            {
                RequireSuccess(await worker.ReleaseAsync(CancellationToken.None));
                await run.WaitAsync(TimeSpan.FromSeconds(10));
            });
            var failures = new[] { failure, completion, shutdown }.OfType<Exception>().ToArray();
            if (failures.Length > 0)
                throw new AggregateException("Private-worker cancellation, completion, or cleanup failed.", failures);
            await RunWithCleanupAsync(async (verification, _, reopenedSessions) =>
            {
                var reopened = await OpenSessionAsync(verification, path);
                reopenedSessions[reopened] = 0;
                Assert.Equal("Saved only by the requested running save", await ReadMarkerAsync(verification, reopened));
            });
        });
    }

    [Fact]
    public async Task Pipe_CancelledSave_RevealsOwnedExcelAndRejectsFollowUpUntilCompletion()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "pipe-cancelled-save.xlsx");
            var session = await CreateSessionAsync(service, path);
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, "Preserved after pipe cancellation");
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(session));
            var identity = Assert.Single(SessionManager.GetTrackedExcelProcesses(),
                candidate => candidate.ProcessId == batch.ExcelProcessId);
            using var release = new ManualResetEventSlim();
            using var releaseSave = new ManualResetEventSlim();
            var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var saving = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            batch.BeforeRefreshStateReadHookForTests = () =>
            {
                entered.TrySetResult();
                Assert.True(release.Wait(TimeSpan.FromSeconds(90)));
            };
            ExcelShutdownService.BeforeSaveHookForTests = () =>
            {
                saving.TrySetResult();
                Assert.True(releaseSave.Wait(TimeSpan.FromSeconds(90)));
            };
            var pipe = $"excelmcp-cancelled-save-{Guid.NewGuid():N}";
            var host = service.RunAsync(pipe);
            using var client = new ServiceClient(pipe);
            using var cancellation = new CancellationTokenSource();
            var save = client.SendAsync(new ServiceRequest
            {
                Command = "session.save",
                SessionId = session
            }, cancellation.Token);
            var failure = await Record.ExceptionAsync(async () =>
            {
                await entered.Task.WaitAsync(TimeSpan.FromSeconds(15));
                var startedAt = batch.OperationStatus.StartedAtUtc;
                await AssertPipeStatusAsync(waitExpired: false, executionPending: true, startedAt);
                cancellation.Cancel();
                await Assert.ThrowsAnyAsync<OperationCanceledException>(
                    () => save.WaitAsync(TimeSpan.FromSeconds(5)));
                Assert.True(SpinWait.SpinUntil(() => batch.OperationStatus.WaitExpired
                    && batch.OperationStatus.WindowReveal is "shown" or "showRequested" or "alreadyVisible",
                    TimeSpan.FromSeconds(5)));
                Assert.True(batch.OperationStatus.ExecutionPending);
                await AssertPipeStatusAsync(waitExpired: true, executionPending: true, startedAt);
                Assert.True(OwnedProcessGuard.MatchesLiveProcess(identity));
                Assert.Same(batch, service.SessionManager.GetSession(session));
                var followUpRequest = new ServiceRequest
                {
                    Command = "range.set-values",
                    SessionId = session,
                    Args = """{"sheetName":"Sheet1","rangeAddress":"A1","overwritePolicy":"allow","values":[["Must not execute"]]}"""
                };
                Assert.False((await client.SendAsync(followUpRequest)
                    .WaitAsync(TimeSpan.FromSeconds(5))).Success);
                batch.BeforeRefreshStateReadHookForTests = null;
                release.Set();
                await saving.Task.WaitAsync(TimeSpan.FromSeconds(15));
                Assert.True(batch.OperationStatus.ExecutionPending);
                Assert.True(batch.HasTimedOutOperation);
                Assert.False((await client.SendAsync(followUpRequest)
                    .WaitAsync(TimeSpan.FromSeconds(5))).Success);
            });
            batch.BeforeRefreshStateReadHookForTests = null;
            ExcelShutdownService.BeforeSaveHookForTests = null;
            release.Set();
            releaseSave.Set();
            var completion = await Record.ExceptionAsync(async () =>
            {
                Assert.True(SpinWait.SpinUntil(() => !batch.OperationStatus.ExecutionPending,
                    TimeSpan.FromSeconds(15)));
                await AssertPipeStatusAsync(waitExpired: true, executionPending: false,
                    batch.OperationStatus.StartedAtUtc);
                Assert.Same(batch, service.SessionManager.GetSession(session));
                Assert.Equal("Preserved after pipe cancellation", await ReadMarkerAsync(service, session));
                Assert.True(batch.IsExcelVisible);
                await CloseSessionAsync(service, session, save: false);
                sessions.TryRemove(session, out _);
                var reopened = await OpenSessionAsync(service, path);
                sessions[reopened] = 0;
                Assert.Equal("Preserved after pipe cancellation", await ReadMarkerAsync(service, reopened));
                await CloseSessionAsync(service, reopened, save: false);
                sessions.TryRemove(reopened, out _);
            });
            var shutdown = await Record.ExceptionAsync(async () =>
            {
                service.RequestShutdown();
                await host.WaitAsync(TimeSpan.FromSeconds(10));
            });
            var failures = new[] { failure, completion, shutdown }.OfType<Exception>().ToArray();
            if (failures.Length > 0)
                throw new AggregateException("Pipe cancellation, retained execution, or shutdown failed.", failures);

            async Task AssertPipeStatusAsync(bool waitExpired, bool executionPending, DateTime startedAt)
            {
                foreach (var command in StatusCommands)
                {
                    using var statusClient = new ServiceClient(pipe);
                    var response = await statusClient.SendAsync(new ServiceRequest { Command = command })
                        .WaitAsync(TimeSpan.FromSeconds(2));
                    RequireSuccess(response);
                    if (command != "session.list")
                        continue;
                    using var result = JsonDocument.Parse(response.Result!);
                    var entry = Assert.Single(result.RootElement.GetProperty("sessions").EnumerateArray());
                    Assert.Equal(session, entry.GetProperty("sessionId").GetString());
                    var operation = entry.GetProperty("operation");
                    Assert.Equal("session.save", operation.GetProperty("command").GetString());
                    Assert.Equal(startedAt, operation.GetProperty("startedAtUtc").GetDateTime());
                    Assert.Equal(waitExpired, operation.GetProperty("waitExpired").GetBoolean());
                    Assert.Equal(executionPending, operation.GetProperty("executionPending").GetBoolean());
                    Assert.Equal(executionPending ? "running" : "completed",
                        operation.GetProperty("state").GetString());
                }
            }
        });
    }

    [Fact]
    public async Task Pipe_CancelledQueuedMutation_DoesNotExecuteAfterReadinessScan()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var session = await CreateSessionAsync(service, Path.Join(directory, "cancelled-queue.xlsx"));
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, "Original contents");
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(session));
            using var release = new ManualResetEventSlim();
            var scanning = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var queued = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            batch.BeforeRefreshStateReadHookForTests = () =>
            {
                scanning.TrySetResult();
                Assert.True(release.Wait(TimeSpan.FromSeconds(90)));
            };
            var inspection = Task.Run(() => service.SessionManager.ValidateClose(session));
            var pipe = $"excelmcp-cancelled-queue-{Guid.NewGuid():N}";
            var host = service.RunAsync(pipe);
            using var client = new ServiceClient(pipe);
            using var cancellation = new CancellationTokenSource();
            var failure = await Record.ExceptionAsync(async () =>
            {
                await scanning.Task.WaitAsync(TimeSpan.FromSeconds(15));
                Assert.False((await inspection.WaitAsync(TimeSpan.FromSeconds(5))).CanClose);
                Assert.Equal("completed", batch.OperationStatus.State);
                ExcelBatch.WorkItemQueuedHookForTests = () => queued.TrySetResult();
                var mutation = client.SendAsync(new ServiceRequest
                {
                    Command = "range.set-values",
                    SessionId = session,
                    Args = """{"sheetName":"Sheet1","rangeAddress":"A1","overwritePolicy":"allow","values":[["Must not execute"]]}"""
                }, cancellation.Token);
                await queued.Task.WaitAsync(TimeSpan.FromSeconds(15));
                cancellation.Cancel();
                await Assert.ThrowsAnyAsync<OperationCanceledException>(
                    () => mutation.WaitAsync(TimeSpan.FromSeconds(5)));
                Assert.True(SpinWait.SpinUntil(() =>
                    batch.OperationStatus.WindowReveal is "shown" or "showRequested" or "alreadyVisible",
                    TimeSpan.FromSeconds(5)));
            });
            ExcelBatch.WorkItemQueuedHookForTests = null;
            batch.BeforeRefreshStateReadHookForTests = null;
            release.Set();
            var completion = await Record.ExceptionAsync(async () =>
            {
                Assert.True(SpinWait.SpinUntil(
                    () => service.SessionManager.GetActiveOperationCount(session) == 0,
                    TimeSpan.FromSeconds(15)));
                Assert.Equal("Original contents", await ReadMarkerAsync(service, session));
                Assert.Same(batch, service.SessionManager.GetSession(session));
                Assert.True(batch.IsExcelProcessAlive());
                await CloseSessionAsync(service, session, save: false);
                sessions.TryRemove(session, out _);
            });
            var shutdown = await Record.ExceptionAsync(async () =>
            {
                service.RequestShutdown();
                await host.WaitAsync(TimeSpan.FromSeconds(10));
            });
            var failures = new[] { failure, completion, shutdown }.OfType<Exception>().ToArray();
            if (failures.Length > 0)
                throw new AggregateException("Queued cancellation, original workbook state, or shutdown failed.", failures);
        });
    }

    [Fact]
    public async Task Test_InvalidWorkbook_PendingQuitRetainsOwnershipDuringServiceShutdown()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "invalid-pending-quit.xlsx");
            await File.WriteAllTextAsync(path, "not an Excel workbook");
            using var release = new ManualResetEventSlim();
            var quitting = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var shutdownWait = ProcessTerminationPolicy.NormalGraceTimeout + TimeSpan.FromSeconds(15);
            ExcelShutdownService.BeforeQuitHookForTests = () =>
            {
                quitting.TrySetResult();
                Assert.True(release.Wait(TimeSpan.FromSeconds(90)));
            };
            var validation = Task.Run(() => service.ProcessAsync(new ServiceRequest
            {
                Command = "session.test",
                Args = JsonSerializer.Serialize(new { filePath = path }, ServiceProtocol.JsonOptions)
            }));
            ExcelBatch? batch = null;
            ExcelProcessIdentity? identity = null;
            var failure = await Record.ExceptionAsync(async () =>
            {
                await quitting.Task.WaitAsync(TimeSpan.FromSeconds(15));
                var response = await validation.WaitAsync(shutdownWait);
                Assert.False(response.Success);
                var session = Assert.IsType<string>(response.SessionId);
                sessions[session] = 0;
                batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(session));
                identity = Assert.Single(SessionManager.GetTrackedExcelProcesses(),
                    candidate => candidate.ProcessId == batch.ExcelProcessId);
                Assert.Contains("cannot open", response.InnerError, StringComparison.OrdinalIgnoreCase);

                var report = service.RequestShutdown();
                Assert.Contains($"Excel process {batch.ExcelProcessId}", report, StringComparison.Ordinal);
                service.Dispose();
                Assert.Same(batch, service.SessionManager.GetSession(session));
                Assert.True(batch.OperationStatus.ExecutionPending);
                Assert.True(OwnedProcessGuard.MatchesLiveProcess(identity.Value));
                Assert.True(service.SessionManager.TryGetFilePath(session, out var retainedPath));
                Assert.Equal(path, retainedPath);
                Assert.Equal("not an Excel workbook", await File.ReadAllTextAsync(path));
            });
            ExcelShutdownService.BeforeQuitHookForTests = null;
            release.Set();
            var cleanup = await Record.ExceptionAsync(async () =>
            {
                await validation.WaitAsync(shutdownWait);
                if (batch is not null)
                    Assert.True(SpinWait.SpinUntil(() => !batch.OperationStatus.ExecutionPending,
                        shutdownWait));
                if (identity is { } owned)
                    Assert.True(SpinWait.SpinUntil(() => OwnedProcessGuard.TryConfirmExited(owned),
                        shutdownWait));
            });
            var failures = new[] { failure, cleanup }.OfType<Exception>().ToArray();
            if (failures.Length > 0)
                throw new AggregateException("Failed-startup shutdown ownership or completion failed.", failures);
        });
    }

    [Theory]
    [InlineData(ResiliencePipelines.RPC_E_SERVERCALL_RETRYLATER)]
    [InlineData(ResiliencePipelines.RPC_E_CALL_REJECTED)]
    public async Task Close_TransientQuitFailure_ReleasesSessionAfterSuccessfulRetry(int hresult)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "quit-retry.xlsx");
            var session = await CreateSessionAsync(service, path);
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, "Saved before retrying quit");
            var batch = service.SessionManager.GetSession(session)!;
            var identity = Assert.Single(SessionManager.GetTrackedExcelProcesses(),
                candidate => candidate.ProcessId == batch.ExcelProcessId);
            var attempts = 0;
            ExcelShutdownService.BeforeQuitHookForTests = () =>
            {
                if (Interlocked.Increment(ref attempts) == 1)
#pragma warning disable CA2201 // A real COMException exercises the production Quit HRESULT retry policy.
                    throw new COMException("Transient quit failure.", hresult);
#pragma warning restore CA2201
            };
            try
            {
                await CloseSessionAsync(service, session, save: true);
            }
            finally
            {
                ExcelShutdownService.BeforeQuitHookForTests = null;
            }

            Assert.Equal(2, attempts);
            Assert.Empty(service.SessionManager.GetActiveSessions());
            Assert.False(service.SessionManager.TryGetFilePath(session, out _));
            Assert.True(OwnedProcessGuard.TryConfirmExited(identity));
            Assert.DoesNotContain(identity, SessionManager.GetTrackedExcelProcesses());
            sessions.TryRemove(session, out _);
            var reopened = await OpenSessionAsync(service, path);
            sessions[reopened] = 0;
            Assert.Equal("Saved before retrying quit", await ReadMarkerAsync(service, reopened));
        });
    }

    [Fact]
    public async Task Test_InvalidWorkbook_CompletesTeardownWithoutRetainingSession()
    {
        await RunWithCleanupAsync(async (service, directory, _) =>
        {
            var path = Path.Join(directory, "invalid-workbook.xlsx");
            await File.WriteAllTextAsync(path, "not an Excel workbook");
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.test",
                Args = JsonSerializer.Serialize(new { filePath = path }, ServiceProtocol.JsonOptions)
            });

            Assert.True(response.Success, $"{response.ErrorMessage} {response.InnerError}");
            using var result = JsonDocument.Parse(response.Result!);
            Assert.False(result.RootElement.GetProperty("isValid").GetBoolean());
            Assert.False(result.RootElement.GetProperty("canOpen").GetBoolean());
            Assert.Equal(path, result.RootElement.GetProperty("filePath").GetString());
            Assert.Contains("cannot open", result.RootElement.GetProperty("message").GetString(),
                StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, service.SessionCount);
            Assert.Empty(service.SessionManager.GetActiveSessions());
            Assert.Equal("not an Excel workbook", await File.ReadAllTextAsync(path));
        });
    }

    [Fact]
    public async Task Test_CloseWaitExpires_ReturnsRetainedSessionInsteadOfInvalidWorkbook()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "validation-teardown.xlsx");
            var initial = await CreateSessionAsync(service, path);
            await WriteMarkerAsync(service, initial, "Valid workbook contents");
            await CloseSessionAsync(service, initial, save: true);
            using var release = new ManualResetEventSlim();
            var cleaning = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            ExcelBatch? batch = null;
            string? session = null;
            ExcelProcessIdentity? identity = null;
            ExcelBatch.AfterWorkbookOpenHookForTests = (_, workbook) =>
            {
                Assert.True(((Excel.Workbook)workbook).ReadOnly);
                session = Assert.Single(service.SessionManager.GetActiveSessions()).SessionId;
                sessions[session] = 0;
                batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(session));
                identity = Assert.Single(SessionManager.GetTrackedExcelProcesses(),
                    candidate => candidate.ProcessId == batch.ExcelProcessId);
                batch.BeforeStaCleanupHookForTests = () =>
                {
                    cleaning.TrySetResult();
                    Assert.True(release.Wait(TimeSpan.FromSeconds(90)));
                };
            };
            var validation = Task.Run(() => service.ProcessAsync(new ServiceRequest
            {
                Command = "session.test",
                Args = JsonSerializer.Serialize(new { filePath = path }, ServiceProtocol.JsonOptions)
            }));
            var failure = await Record.ExceptionAsync(async () =>
            {
                await cleaning.Task.WaitAsync(TimeSpan.FromSeconds(15));
                var response = await validation.WaitAsync(ComInteropConstants.StaThreadJoinTimeout + TimeSpan.FromSeconds(10));
                Assert.False(response.Success);
                Assert.Equal("Timeout", response.ErrorCategory);
                Assert.Equal("session.test", response.Command);
                Assert.Equal(session, response.SessionId);
                Assert.Contains("close/quit", response.InnerError, StringComparison.Ordinal);
                using var partial = JsonDocument.Parse(Assert.IsType<string>(response.PartialState));
                Assert.True(partial.RootElement.GetProperty("operation").GetProperty("executionPending").GetBoolean());
                Assert.True(partial.RootElement.GetProperty("operation").GetProperty("waitExpired").GetBoolean());
                Assert.Same(batch, service.SessionManager.GetSession(session!));
                Assert.True(OwnedProcessGuard.MatchesLiveProcess(identity!.Value));
                Assert.True(service.SessionManager.TryGetFilePath(session!, out var retainedPath));
                Assert.Equal(path, retainedPath);
            });
            ExcelBatch.AfterWorkbookOpenHookForTests = null;
            release.Set();
            var cleanup = await Record.ExceptionAsync(async () =>
            {
                await validation.WaitAsync(TimeSpan.FromSeconds(20));
                if (identity is { } owned)
                    Assert.True(SpinWait.SpinUntil(() => OwnedProcessGuard.TryConfirmExited(owned), TimeSpan.FromSeconds(15)));
                if (session is not null)
                {
                    await CloseSessionAsync(service, session, save: false);
                    sessions.TryRemove(session, out _);
                }
            });
            var failures = new[] { failure, cleanup }.OfType<Exception>().ToArray();
            if (failures.Length > 0)
                throw new AggregateException("Validation teardown timeout or retained cleanup failed.", failures);
            var reopened = await OpenSessionAsync(service, path);
            sessions[reopened] = 0;
            Assert.Equal("Valid workbook contents", await ReadMarkerAsync(service, reopened));
        });
    }

    [Fact]
    public async Task Open_LateFailure_TracksPendingCleanupAndPreservesError()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "late-startup-failure.xlsx");
            var initial = await CreateSessionAsync(service, path);
            await CloseSessionAsync(service, initial, save: true);
            using var finishOpen = new ManualResetEventSlim();
            using var finishCleanup = new ManualResetEventSlim();
            var opening = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var cleaning = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            ExcelBatch.AfterWorkbookOpenHookForTests = (_, workbook) =>
            {
                Assert.Equal(path, ((Excel.Workbook)workbook).FullName, ignoreCase: true);
                opening.TrySetResult();
                Assert.True(finishOpen.Wait(TimeSpan.FromSeconds(30)));
                throw new InvalidOperationException("Original late startup failure.");
            };
            var creation = Task.Run(() => service.SessionManager.CreateSessionWithTimeouts(
                path, false, TimeSpan.FromSeconds(2), TimeSpan.FromSeconds(2)));
            ExcelBatch? batch = null;
            string? session = null;
            ExcelProcessIdentity? identity = null;
            var failure = await Record.ExceptionAsync(async () =>
            {
                await opening.Task.WaitAsync(TimeSpan.FromSeconds(15));
                var pending = await Assert.ThrowsAsync<SessionStartupPendingException>(() => creation);
                session = pending.SessionId;
                sessions[session] = 0;
                batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(session));
                identity = Assert.Single(SessionManager.GetTrackedExcelProcesses(),
                    candidate => candidate.ProcessId == batch.ExcelProcessId);
                batch.BeforeStaCleanupHookForTests = () =>
                {
                    cleaning.TrySetResult();
                    Assert.True(finishCleanup.Wait(TimeSpan.FromSeconds(30)));
                };
                finishOpen.Set();
                await cleaning.Task.WaitAsync(TimeSpan.FromSeconds(10));
                var pendingCleanup = service.SessionManager.GetOperationStatus(session)!;
                Assert.True(pendingCleanup.ExecutionPending);
                Assert.True(pendingCleanup.WaitExpired);
                Assert.Equal("Original late startup failure.", pendingCleanup.ErrorMessage);
                Assert.True(batch.IsExcelProcessAlive());
                Assert.Equal(session, Assert.Single(service.SessionManager.GetActiveSessions()).SessionId);
            });
            ExcelBatch.AfterWorkbookOpenHookForTests = null;
            finishOpen.Set();
            finishCleanup.Set();
            var cleanupFailure = await Record.ExceptionAsync(async () =>
            {
                if (batch is null)
                    return;
                Assert.True(SpinWait.SpinUntil(
                    () => !batch.OperationStatus.ExecutionPending,
                    TimeSpan.FromSeconds(15)));
                var completed = batch.OperationStatus;
                Assert.Equal("failed", completed.State);
                Assert.Contains("Original late startup failure.", completed.ErrorMessage, StringComparison.Ordinal);
                Assert.True(SpinWait.SpinUntil(
                    () => OwnedProcessGuard.TryConfirmExited(identity!.Value),
                    TimeSpan.FromSeconds(15)));
            });
            if (session is not null)
            {
                await CloseSessionAsync(service, session, save: false);
                sessions.TryRemove(session, out _);
            }
            var failures = new[] { failure, cleanupFailure }.OfType<Exception>().ToArray();
            if (failures.Length > 0)
                throw new AggregateException("Late startup failure or pending cleanup failed.", failures);
        });
    }

    [Theory]
    [InlineData("workbook.save-copy-as", "xlsx")]
    [InlineData("workbook.export-fixed-format", "pdf")]
    [InlineData("workbook.export-fixed-format", "xps")]
    public async Task Output_WaitExpires_CompletesCommitAndTemporaryCleanupOnOriginalCallback(
        string command, string extension)
    {
        if (extension != "xlsx")
            FixedFormatExportPrerequisite.RequireUsableDefaultPrinter();
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "late-output-source.xlsx");
            var target = Path.Join(directory, $"late-output.{extension}");
            var initial = await CreateSessionAsync(service, path);
            await WriteMarkerAsync(service, initial, "Late output contents");
            await CloseSessionAsync(service, initial, save: true);
            var session = service.SessionManager.CreateSessionWithTimeouts(
                path, false, TimeSpan.FromSeconds(2), ComInteropConstants.DefaultOperationTimeout);
            sessions[session] = 0;
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(session));
            await File.WriteAllTextAsync(target, "Original destination");
            using var release = new ManualResetEventSlim();
            var entered = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
            WorkbookCommands.BeforeOutputCommitHookForTests = temporaryPath =>
            {
                entered.TrySetResult(temporaryPath);
                Assert.True(release.Wait(TimeSpan.FromSeconds(30)));
            };
            var operation = Task.Run(() => service.ProcessAsync(new ServiceRequest
            {
                Command = command,
                SessionId = session,
                Args = command == "workbook.save-copy-as"
                    ? JsonSerializer.Serialize(new { targetPath = target, overwrite = true }, ServiceProtocol.JsonOptions)
                    : JsonSerializer.Serialize(new { targetPath = target, formatType = extension, overwrite = true },
                        ServiceProtocol.JsonOptions)
            }));
            string? temporary = null;
            var failure = await Record.ExceptionAsync(async () =>
            {
                temporary = await entered.Task.WaitAsync(TimeSpan.FromSeconds(15));
                var response = await operation.WaitAsync(TimeSpan.FromSeconds(10));
                Assert.False(response.Success);
                Assert.Equal("Timeout", response.ErrorCategory);
                Assert.Equal(session, response.SessionId);
                using var partial = JsonDocument.Parse(Assert.IsType<string>(response.PartialState));
                Assert.True(partial.RootElement.GetProperty("operation").GetProperty("executionPending").GetBoolean());
                Assert.NotEqual(target, temporary);
                Assert.True(File.Exists(temporary));
                Assert.Equal("Original destination", await File.ReadAllTextAsync(target));
                Assert.True(batch.IsExcelProcessAlive());
            });
            release.Set();
            var completionFailure = await Record.ExceptionAsync(async () =>
            {
                await operation.WaitAsync(TimeSpan.FromSeconds(20));
                Assert.True(SpinWait.SpinUntil(
                    () => batch.GetRefreshState() == WorkbookRefreshState.Ready,
                    TimeSpan.FromSeconds(15)));
                var completed = batch.OperationStatus;
                Assert.Equal("completed", completed.State);
                Assert.True(completed.WaitExpired);
                Assert.False(completed.ExecutionPending);
                Assert.Null(completed.ErrorMessage);
                Assert.NotNull(temporary);
                Assert.False(File.Exists(temporary));
            });
            WorkbookCommands.BeforeOutputCommitHookForTests = null;
            var failures = new[] { failure, completionFailure }.OfType<Exception>().ToArray();
            if (failures.Length > 0)
                throw new AggregateException("Late output commit or cleanup failed.", failures);

            Assert.Equal(path, batch.WorkbookPath, ignoreCase: true);
            Assert.Equal("Late output contents", await ReadMarkerAsync(service, session));
            if (extension == "xlsx")
            {
                var copy = await OpenSessionAsync(service, target);
                sessions[copy] = 0;
                Assert.Equal("Late output contents", await ReadMarkerAsync(service, copy));
            }
            else if (extension == "pdf")
            {
                var bytes = await File.ReadAllBytesAsync(target);
                Assert.True(bytes.Length > 100);
                Assert.Equal("%PDF", System.Text.Encoding.ASCII.GetString(bytes, 0, 4));
            }
            else
            {
                using var package = System.IO.Compression.ZipFile.OpenRead(target);
                var page = Assert.Single(package.Entries,
                    entry => entry.FullName.EndsWith(".fpage", StringComparison.OrdinalIgnoreCase));
                using var pageStream = page.Open();
                var xml = System.Xml.Linq.XDocument.Load(pageStream);
                Assert.Contains("Late output contents", xml.ToString(), StringComparison.Ordinal);
            }
            Assert.Empty(Directory.GetFiles(directory, "*.tmp.*"));
        });
    }

    [Theory]
    [InlineData("session.save")]
    [InlineData("workbook.save-as")]
    public async Task Save_SlowReadinessScan_UsesOperationBudgetAndPersistsContents(string command)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "readiness-save.xlsx");
            var target = command == "session.save" ? path : Path.Join(directory, "readiness-save-as.xlsx");
            var session = await CreateSessionAsync(service, path);
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, "Persisted after readiness");
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(session));
            batch.BeforeRefreshStateReadHookForTests = () => Thread.Sleep(400);
            try
            {
                var response = await service.ProcessAsync(new ServiceRequest
                {
                    Command = command,
                    SessionId = session,
                    Args = command == "workbook.save-as"
                        ? JsonSerializer.Serialize(new { targetPath = target }, ServiceProtocol.JsonOptions)
                        : null
                });
                RequireSuccess(response);
                using var saved = JsonDocument.Parse(response.Result!);
                if (command == "session.save")
                {
                    Assert.True(saved.RootElement.GetProperty("saved").GetBoolean());
                    Assert.Equal(session, saved.RootElement.GetProperty("sessionId").GetString());
                }
                else
                {
                    Assert.True(saved.RootElement.GetProperty("success").GetBoolean());
                    Assert.Equal(target, saved.RootElement.GetProperty("filePath").GetString());
                }
            }
            finally
            {
                batch.BeforeRefreshStateReadHookForTests = null;
            }
            await CloseSessionAsync(service, session, save: false);
            sessions.TryRemove(session, out _);
            var reopened = await OpenSessionAsync(service, target);
            sessions[reopened] = 0;
            Assert.Equal("Persisted after readiness", await ReadMarkerAsync(service, reopened));
        });
    }

    [Theory]
    [InlineData(false, false, false, null)]
    [InlineData(true, false, false, null)]
    [InlineData(false, true, false, null)]
    [InlineData(false, false, true, null)]
    [InlineData(false, false, false, "session.attach")]
    [InlineData(false, false, false, "workbook.save-as")]
    public async Task Controller_LateSaveAsPublishesFinalPathAndPreservesReservations(
        bool cancel, bool completeDuringCapture, bool checkSessionlessRouting, string? rejectedCommand)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "controller-late-save.xlsx");
            var target = Path.Join(directory, "controller-late-save-as.xlsx");
            var unusedTarget = Path.Join(directory, "rejected-save-as.xlsx");
            var neighborPath = Path.Join(directory, "rejected-attachment.xlsx");
            var session = await CreateSessionAsync(service, path);
            await WriteMarkerAsync(service, session, "Saved baseline");
            await CloseSessionAsync(service, session, save: true);
            session = service.SessionManager.CreateSessionWithTimeouts(
                path, true, TimeSpan.FromSeconds(checkSessionlessRouting || rejectedCommand is not null ? 10 : 2),
                ComInteropConstants.DefaultOperationTimeout);
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, "Saved after timeout");
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(session));
            string? neighborName = null;
            string? neighborSession = null;
            if (rejectedCommand == "session.attach")
                neighborName = batch.Execute((context, _) =>
                {
                    var name = AddReviewWorkbook(context.App, "Neighbor retained after Busy");
                    Excel.Workbooks? books = null;
                    Excel.Workbook? workbook = null;
                    try
                    {
                        books = context.App.Workbooks;
                        workbook = books[name];
                        workbook.SaveAs(neighborPath, ComInteropConstants.XlOpenXmlWorkbook);
                        return workbook.Name;
                    }
                    finally
                    {
                        ComUtilities.Release(ref workbook);
                        ComUtilities.Release(ref books);
                    }
                });
            var identity = Assert.IsType<ExcelProcessIdentity>(batch.ProcessIdentity);
            var executable = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
                @"..\..\..\..\..\src\ExcelMcp.CLI\bin\Release\net10.0-windows\excelcli.exe"));
            using var controller = new ExcelController(launcher: new ProductProcessLauncher(executable, []));
            var bootstrap = await controller.StartWorkerAsync(null, CancellationToken.None);
            var registration = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}",
                WorkerProcess.Current, WorkerProcess.Current, bootstrap.Registration.Generation,
                WorkerRegistry.Build, identity, OwnsExcel: true);
            using var host = new ExcelWorkerHost(registration, service);
            var run = host.RunAsync();
            var worker = new ExcelWorkerClient(registration, WorkerProcess.Current, registration.Generation);
            using var release = new ManualResetEventSlim();
            using var resumeCapture = new ManualResetEventSlim();
            var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            Excel.AppEvents_WorkbookBeforeSaveEventHandler beforeSave =
                (Excel.Workbook _, bool _, ref bool cancelSave) =>
                {
                    entered.TrySetResult();
                    Assert.True(release.Wait(TimeSpan.FromSeconds(30)));
                    cancelSave = cancel;
                };
            batch.Execute((ctx, _) => ctx.App.WorkbookBeforeSave += beforeSave);
            var claimed = false;
            Task<ServiceResponse>? saving = null;
            var failure = await Record.ExceptionAsync(async () =>
            {
                await worker.InitializeAsync();
                var discovered = await controller.ProcessAsync(new ServiceRequest { Command = "session.discover" });
                RequireSuccess(discovered);
                using var discovery = JsonDocument.Parse(discovered.Result!);
                Assert.True(discovery.RootElement.GetProperty("complete").GetBoolean(), discovered.Result);
                var selected = Assert.Single(discovery.RootElement.GetProperty("workbooks").EnumerateArray(),
                    item => item.GetProperty("location").GetString() == path);
                var neighborSelection = neighborName is null ? null
                    : Assert.Single(discovery.RootElement.GetProperty("workbooks").EnumerateArray(),
                        item => item.GetProperty("location").GetString() == neighborPath)
                        .GetProperty("workbookId").GetString();
                claimed = true;
                var attached = await controller.ProcessAsync(new ServiceRequest
                {
                    Command = "session.attach",
                    Args = JsonSerializer.Serialize(new { workbookId = selected.GetProperty("workbookId").GetString() })
                });
                RequireSuccess(attached);
                Assert.Equal(session, attached.SessionId);
                saving = controller.ProcessAsync(new ServiceRequest
                {
                    Command = "workbook.save-as",
                    SessionId = session,
                    Args = JsonSerializer.Serialize(new { targetPath = target, overwrite = true })
                });
                await entered.Task.WaitAsync(TimeSpan.FromSeconds(10));
                var rejectedRequest = rejectedCommand is null ? null : new ServiceRequest
                {
                    Command = rejectedCommand,
                    SessionId = rejectedCommand == "workbook.save-as" ? session : null,
                    Args = rejectedCommand == "session.attach"
                        ? JsonSerializer.Serialize(new { workbookId = neighborSelection })
                        : JsonSerializer.Serialize(new { targetPath = unusedTarget })
                };
                if (rejectedRequest is not null)
                {
                    var rejected = await controller.ProcessAsync(rejectedRequest).WaitAsync(TimeSpan.FromSeconds(5));
                    Assert.False(rejected.Success);
                    Assert.Equal("Busy", rejected.ErrorCategory);
                    Assert.True(batch.OperationStatus.ExecutionPending);
                }
                if (checkSessionlessRouting)
                {
                    var echo = await controller.ProcessAsync(new ServiceRequest
                    {
                        Command = "diag.echo",
                        SessionId = session,
                        Args = """{"message":"during-save"}"""
                    }).WaitAsync(TimeSpan.FromSeconds(5));
                    RequireSuccess(echo);
                    using var echoResult = JsonDocument.Parse(echo.Result!);
                    Assert.Equal("during-save", echoResult.RootElement.GetProperty("message").GetString());
                    var settingRequest = new ServiceRequest
                    {
                        Command = "settings.get-native-query-approval",
                        SessionId = session
                    };
                    var expectedSetting = await service.ProcessAsync(settingRequest);
                    RequireSuccess(expectedSetting);
                    var actualSetting = await controller.ProcessAsync(settingRequest).WaitAsync(TimeSpan.FromSeconds(5));
                    RequireSuccess(actualSetting);
                    Assert.Equal(expectedSetting.Result, actualSetting.Result);
                    Assert.False(saving.IsCompleted);
                    Assert.True(batch.OperationStatus.ExecutionPending);
                }
                var timedOut = await saving.WaitAsync(TimeSpan.FromSeconds(15));
                Assert.False(timedOut.Success);
                Assert.Equal("Timeout", timedOut.ErrorCategory);
                if (rejectedRequest is not null)
                {
                    var rejected = await controller.ProcessAsync(rejectedRequest).WaitAsync(TimeSpan.FromSeconds(5));
                    Assert.False(rejected.Success);
                    Assert.Equal("Busy", rejected.ErrorCategory);
                    Assert.True(batch.OperationStatus.ExecutionPending);
                }
                await Task.Delay(1500);
                RequireSuccess(await controller.ProcessAsync(new ServiceRequest
                {
                    Command = "window.show",
                    SessionId = session
                }));
                var pendingTarget = await controller.ProcessAsync(new ServiceRequest
                {
                    Command = "session.open",
                    Args = JsonSerializer.Serialize(new { filePath = target })
                });
                Assert.False(pendingTarget.Success);
                Assert.Equal("Conflict", pendingTarget.ErrorCategory);
                if (completeDuringCapture)
                {
                    var captures = 0;
                    var captureBlocked = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
                    host.AfterOperationStatusCaptureForTests = () =>
                    {
                        if (Interlocked.Increment(ref captures) == 1)
                        {
                            release.Set();
                            Assert.True(SpinWait.SpinUntil(
                                () => !batch.OperationStatus.ExecutionPending, TimeSpan.FromSeconds(15)));
                        }
                        else
                        {
                            captureBlocked.TrySetResult();
                            Assert.True(resumeCapture.Wait(TimeSpan.FromSeconds(15)));
                        }
                    };
                    _ = await worker.ReadSnapshotAsync(CancellationToken.None);
                    await captureBlocked.Task.WaitAsync(TimeSpan.FromSeconds(10));
                    await Task.Delay(1500);
                    var settledTarget = await controller.ProcessAsync(new ServiceRequest
                    {
                        Command = "session.create",
                        Args = JsonSerializer.Serialize(new { filePath = target })
                    });
                    Assert.False(settledTarget.Success);
                    Assert.Equal("Conflict", settledTarget.ErrorCategory);
                    Assert.Equal("This workbook is already reserved by another session or pending request.",
                        settledTarget.ErrorMessage);
                    host.AfterOperationStatusCaptureForTests = null;
                    resumeCapture.Set();
                }
                release.Set();
                var timer = Stopwatch.StartNew();
                while (batch.GetRefreshState() != WorkbookRefreshState.Ready)
                {
                    Assert.True(timer.Elapsed < TimeSpan.FromSeconds(15));
                    await Task.Delay(100);
                }
                var expectedPath = cancel ? path : target;
                timer.Restart();
                while (true)
                {
                    var status = await controller.ProcessAsync(new ServiceRequest { Command = "session.list" });
                    RequireSuccess(status);
                    using var statusList = JsonDocument.Parse(status.Result!);
                    var statusEntry = Assert.Single(statusList.RootElement.GetProperty("sessions").EnumerateArray());
                    if (Assert.Single(controller.Sessions).FilePath == expectedPath
                        && statusEntry.GetProperty("filePath").GetString() == expectedPath
                        && !statusEntry.GetProperty("operationPending").GetBoolean())
                        break;
                    Assert.True(timer.Elapsed < TimeSpan.FromSeconds(10), "The controller kept the old Save As path.");
                    await Task.Delay(100);
                }
                var listed = await controller.ProcessAsync(new ServiceRequest { Command = "session.list" });
                RequireSuccess(listed);
                using var list = JsonDocument.Parse(listed.Result!);
                Assert.Equal(expectedPath, Assert.Single(list.RootElement.GetProperty("sessions").EnumerateArray())
                    .GetProperty("filePath").GetString());
                if (rejectedCommand is not null)
                {
                    var retried = await controller.ProcessAsync(new ServiceRequest
                    {
                        Command = rejectedCommand == "session.attach" ? "session.attach" : "session.create",
                        Args = rejectedCommand == "session.attach"
                            ? JsonSerializer.Serialize(new { workbookId = neighborSelection })
                            : JsonSerializer.Serialize(new { filePath = unusedTarget })
                    });
                    RequireSuccess(retried);
                    if (rejectedCommand == "session.attach")
                    {
                        neighborSession = retried.SessionId;
                        Assert.Equal("Neighbor retained after Busy", await ReadMarkerAsync(service, neighborSession!));
                    }
                    RequireSuccess(await controller.ProcessAsync(new ServiceRequest
                    {
                        Command = "session.close",
                        SessionId = retried.SessionId
                    }));
                    if (rejectedCommand == "session.attach")
                        neighborName = null;
                }
                var availablePath = cancel ? target : path;
                var available = await controller.ProcessAsync(new ServiceRequest
                {
                    Command = cancel ? "session.create" : "session.open",
                    Args = JsonSerializer.Serialize(new { filePath = availablePath })
                });
                RequireSuccess(available);
                RequireSuccess(await controller.ProcessAsync(new ServiceRequest
                {
                    Command = "session.close",
                    SessionId = available.SessionId
                }));
                Assert.Equal("Saved after timeout", await ReadMarkerAsync(service, session));
            });
            release.Set();
            host.AfterOperationStatusCaptureForTests = null;
            resumeCapture.Set();
            var cleanup = await Record.ExceptionAsync(async () =>
            {
                if (saving is not null)
                    _ = await saving.WaitAsync(TimeSpan.FromSeconds(15));
                var timer = Stopwatch.StartNew();
                while (batch.OperationStatus.ExecutionPending)
                {
                    Assert.True(timer.Elapsed < TimeSpan.FromSeconds(15));
                    await Task.Delay(100);
                }
                timer.Restart();
                while (batch.GetRefreshState() != WorkbookRefreshState.Ready)
                {
                    Assert.True(timer.Elapsed < TimeSpan.FromSeconds(15));
                    await Task.Delay(100);
                }
                if (neighborName is not null)
                {
                    if (neighborSession is not null)
                        await CloseSessionAsync(service, neighborSession, save: false);
                    else
                        batch.Execute((context, _) =>
                        {
                            Excel.Workbooks? books = null;
                            Excel.Workbook? workbook = null;
                            try
                            {
                                books = context.App.Workbooks;
                                workbook = books[neighborName];
                                workbook.Close(SaveChanges: false);
                            }
                            finally
                            {
                                ComUtilities.Release(ref workbook);
                                ComUtilities.Release(ref books);
                            }
                        });
                }
                batch.Execute((ctx, _) => ctx.App.WorkbookBeforeSave -= beforeSave);
                await CloseSessionAsync(service, session, save: false);
                sessions.TryRemove(session, out _);
                if (!claimed)
                    RequireSuccess(await worker.ReleaseAsync(CancellationToken.None));
                await controller.RequestShutdownAsync();
                await run.WaitAsync(TimeSpan.FromSeconds(10));
            });
            var failures = new[] { failure, cleanup }.OfType<Exception>().ToArray();
            if (failures.Length > 0) throw new AggregateException("Late Save As or cleanup failed.", failures);
            await RunWithCleanupAsync(async (verification, _, reopenedSessions) =>
            {
                var reopened = await OpenSessionAsync(verification, cancel ? path : target);
                reopenedSessions[reopened] = 0;
                Assert.Equal(cancel ? "Saved baseline" : "Saved after timeout", await ReadMarkerAsync(verification, reopened));
            });
        });
    }

    [Theory]
    [InlineData("session.save", false)]
    [InlineData("workbook.save-as", false)]
    [InlineData("workbook.save-as", true)]
    public async Task Save_WaitExpires_RetainsSessionAndTracksLatePersistence(string command, bool cancel)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "late-save.xlsx");
            var target = Path.Join(directory, "late-save-as.xlsx");
            var session = await CreateSessionAsync(service, path);
            await WriteMarkerAsync(service, session, "Baseline");
            await CloseSessionAsync(service, session, save: true);
            session = service.SessionManager.CreateSessionWithTimeouts(
                path, false, TimeSpan.FromSeconds(2), ComInteropConstants.DefaultOperationTimeout);
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, "Unsaved checkpoint");
            var neighbor = await CreateSessionAsync(service, Path.Join(directory, "neighbor.xlsx"));
            sessions[neighbor] = 0;
            await WriteMarkerAsync(service, neighbor, "Neighbor");
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(session));
            var pid = batch.ExcelProcessId;
            using var release = new ManualResetEventSlim();
            var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            Excel.AppEvents_WorkbookBeforeSaveEventHandler beforeSave =
                (Excel.Workbook _, bool _, ref bool cancelSave) =>
                {
                    entered.TrySetResult();
                    Assert.True(release.Wait(TimeSpan.FromSeconds(30)));
                    cancelSave = cancel;
                };
            batch.Execute((ctx, _) => ctx.App.WorkbookBeforeSave += beforeSave);
            var failure = await Record.ExceptionAsync(async () =>
            {
                var save = Task.Run(() => service.ProcessAsync(new ServiceRequest
                {
                    Command = command,
                    SessionId = session,
                    Args = command == "workbook.save-as"
                        ? JsonSerializer.Serialize(new { targetPath = target, overwrite = true }, ServiceProtocol.JsonOptions)
                        : null
                }));
                await entered.Task.WaitAsync(TimeSpan.FromSeconds(10));
                var response = await save.WaitAsync(TimeSpan.FromSeconds(10));
                Assert.False(response.Success);
                Assert.Equal("Timeout", response.ErrorCategory);
                Assert.Equal(session, response.SessionId);
                Assert.Equal(command, response.Command);
                using var partial = JsonDocument.Parse(Assert.IsType<string>(response.PartialState));
                Assert.True(partial.RootElement.GetProperty("operation").GetProperty("waitExpired").GetBoolean());
                Assert.True(partial.RootElement.GetProperty("operation").GetProperty("executionPending").GetBoolean());
                Assert.Same(batch, service.SessionManager.GetSession(session));
                Assert.Equal(pid, batch.ExcelProcessId);
                Assert.True(batch.IsExcelProcessAlive());

                var timer = Stopwatch.StartNew();
                var list = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
                RequireSuccess(list);
                Assert.True(timer.Elapsed < TimeSpan.FromSeconds(2));
                using var pending = JsonDocument.Parse(list.Result!);
                var entry = Assert.Single(pending.RootElement.GetProperty("sessions").EnumerateArray(),
                    item => item.GetProperty("sessionId").GetString() == session);
                Assert.True(entry.GetProperty("operationPending").GetBoolean());
                Assert.False(entry.GetProperty("canClose").GetBoolean());
                Assert.Equal(command, entry.GetProperty("operation").GetProperty("command").GetString());
                Assert.Equal("Neighbor", await ReadMarkerAsync(service, neighbor));
                var close = await service.ProcessAsync(new ServiceRequest { Command = "session.close", SessionId = session });
                Assert.False(close.Success);
                Assert.Same(batch, service.SessionManager.GetSession(session));
            });
            release.Set();
            var cleanupFailure = await Record.ExceptionAsync(async () =>
            {
                var timer = Stopwatch.StartNew();
                while (batch.GetRefreshState() != WorkbookRefreshState.Ready)
                {
                    Assert.True(timer.Elapsed < TimeSpan.FromSeconds(15));
                    await Task.Delay(100);
                }
                var completed = service.SessionManager.GetOperationStatus(session)!;
                Assert.Equal(cancel ? "failed" : "completed", completed.State);
                if (cancel)
                    Assert.Contains("not saved", completed.ErrorMessage, StringComparison.OrdinalIgnoreCase);
                Assert.True(completed.WaitExpired);
                Assert.False(completed.ExecutionPending);
                batch.Execute((ctx, _) => ctx.App.WorkbookBeforeSave -= beforeSave);
            });
            if (failure is not null && cleanupFailure is not null)
                throw new AggregateException(failure, cleanupFailure);
            if (failure is not null || cleanupFailure is not null)
                ExceptionDispatchInfo.Capture((failure ?? cleanupFailure)!).Throw();

            var expectedPath = command == "workbook.save-as" && !cancel ? target : path;
            Assert.Equal(expectedPath, batch.WorkbookPath, ignoreCase: true);
            Assert.True(service.SessionManager.TryGetFilePath(session, out var trackedPath));
            Assert.Equal(expectedPath, trackedPath, ignoreCase: true);
            Assert.Equal("Unsaved checkpoint", await ReadMarkerAsync(service, session));
            await CloseSessionAsync(service, session, save: false);
            sessions.TryRemove(session, out _);
            var reopened = await OpenSessionAsync(service, expectedPath);
            sessions[reopened] = 0;
            Assert.Equal(cancel ? "Baseline" : "Unsaved checkpoint", await ReadMarkerAsync(service, reopened));
            if (cancel)
            {
                var releasedTarget = await CreateSessionAsync(service, target);
                sessions[releasedTarget] = 0;
                await WriteMarkerAsync(service, releasedTarget, "Released target");
                Assert.Equal("Released target", await ReadMarkerAsync(service, releasedTarget));
            }
            Assert.Equal("Neighbor", await ReadMarkerAsync(service, neighbor));
        });
    }
}
