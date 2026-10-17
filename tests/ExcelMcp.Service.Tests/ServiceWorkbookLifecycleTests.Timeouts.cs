using System.Diagnostics;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Workbook;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    [Fact]
    public async Task Close_CancelledBeforeExecution_OrdinaryShutdownStillSavesUnsavedChanges()
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
            Assert.True(OwnedProcessGuard.TryConfirmExited(identity));
            await RunWithCleanupAsync(async (verification, _, reopenedSessions) =>
            {
                var reopened = await OpenSessionAsync(verification, path);
                reopenedSessions[reopened] = 0;
                Assert.Equal("Latest unsaved changes", await ReadMarkerAsync(verification, reopened));
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
                cancellation.Cancel();
                await Assert.ThrowsAnyAsync<OperationCanceledException>(
                    () => save.WaitAsync(TimeSpan.FromSeconds(5)));
                Assert.True(SpinWait.SpinUntil(() => batch.OperationStatus.WaitExpired
                    && batch.OperationStatus.WindowReveal is "shown" or "showRequested" or "alreadyVisible",
                    TimeSpan.FromSeconds(5)));
                Assert.True(batch.OperationStatus.ExecutionPending);
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
                var response = await validation.WaitAsync(TimeSpan.FromSeconds(30));
                Assert.False(response.Success);
                var session = Assert.IsType<string>(response.SessionId);
                sessions[session] = 0;
                batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(session));
                identity = Assert.Single(SessionManager.GetTrackedExcelProcesses(),
                    candidate => candidate.ProcessId == batch.ExcelProcessId);
                Assert.Contains("cannot open", response.InnerError, StringComparison.OrdinalIgnoreCase);

                var shutdown = Assert.Throws<AggregateException>(service.Dispose);
                Assert.Contains(shutdown.InnerExceptions, ex => ex.InnerException is ExcelBusyException);
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
                await validation.WaitAsync(TimeSpan.FromSeconds(30));
                if (batch is not null)
                    Assert.True(SpinWait.SpinUntil(() => !batch.OperationStatus.ExecutionPending,
                        TimeSpan.FromSeconds(15)));
                if (identity is { } owned)
                    Assert.True(SpinWait.SpinUntil(() => OwnedProcessGuard.TryConfirmExited(owned),
                        TimeSpan.FromSeconds(15)));
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

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Open_LateFailure_TracksPendingCleanupAndPreservesBothErrors(bool cleanupFails)
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
                    if (cleanupFails)
                        throw new InvalidOperationException("Original late cleanup failure.");
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
                if (cleanupFails)
                {
                    Assert.Contains("Original late cleanup failure.", completed.ErrorMessage, StringComparison.Ordinal);
                    Assert.True(OwnedProcessGuard.MatchesLiveProcess(identity!.Value));
                    Assert.Throws<InvalidOperationException>(() => service.SessionManager.CreateSession(path));
                }
                else
                {
                    Assert.True(SpinWait.SpinUntil(
                        () => OwnedProcessGuard.TryConfirmExited(identity!.Value),
                        TimeSpan.FromSeconds(15)));
                }
            });
            var processCleanup = Record.Exception(() =>
            {
                if (cleanupFails && identity is { } owned)
                    Assert.True(OwnedProcessGuard.TryTerminate(owned, TimeSpan.Zero, TimeSpan.FromSeconds(5), out _));
            });
            if (session is not null)
            {
                await CloseSessionAsync(service, session, save: false);
                sessions.TryRemove(session, out _);
            }
            var failures = new[] { failure, cleanupFailure, processCleanup }.OfType<Exception>().ToArray();
            if (failures.Length > 0)
                throw new AggregateException("Late startup failure or retained cleanup failed.", failures);
        });
    }

    [Theory]
    [InlineData("workbook.save-copy-as", "xlsx")]
    [InlineData("workbook.export-fixed-format", "pdf")]
    [InlineData("workbook.export-fixed-format", "xps")]
    public async Task Output_WaitExpires_CompletesCommitAndTemporaryCleanupOnOriginalCallback(
        string command, string extension)
    {
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

    [Fact]
    public async Task Open_StartupAndTeardownFail_RetainsSessionAndBothFailures()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "failed-startup.xlsx");
            var initial = await CreateSessionAsync(service, path);
            await WriteMarkerAsync(service, initial, "Original workbook");
            await CloseSessionAsync(service, initial, save: true);
            ExcelProcessIdentity? identity = null;
            string? retainedSession = null;
            Excel.AppEvents_WorkbookBeforeCloseEventHandler preventClose =
                (Excel.Workbook _, ref bool cancel) => cancel = true;
            ExcelBatch.AfterWorkbookOpenHookForTests = (application, _) =>
            {
                var app = (Excel.Application)application;
                app.WorkbookBeforeClose += preventClose;
                identity = Assert.Single(SessionManager.GetTrackedExcelProcesses());
                throw new InvalidOperationException("Failure after opening the workbook.");
            };
            try
            {
                var response = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.open",
                    Args = JsonSerializer.Serialize(new { filePath = path }, ServiceProtocol.JsonOptions)
                });
                Assert.False(response.Success);
                Assert.False(string.IsNullOrWhiteSpace(response.SessionId));
                retainedSession = response.SessionId;
                sessions[response.SessionId!] = 0;
                Assert.Contains("Failure after opening the workbook", response.InnerError, StringComparison.Ordinal);
                Assert.Contains("close/quit", response.InnerError, StringComparison.Ordinal);
                Assert.NotNull(identity);
                Assert.True(OwnedProcessGuard.MatchesLiveProcess(identity.Value));
                Assert.NotNull(service.SessionManager.GetSession(response.SessionId!));
                Assert.Equal(response.SessionId, Assert.Single(service.SessionManager.GetActiveSessions()).SessionId);
                var repeated = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.open",
                    Args = JsonSerializer.Serialize(new { filePath = path }, ServiceProtocol.JsonOptions)
                });
                Assert.False(repeated.Success);
                Assert.Contains("already open", repeated.ErrorMessage, StringComparison.Ordinal);
            }
            finally
            {
                ExcelBatch.AfterWorkbookOpenHookForTests = null;
                // This fixture intentionally cancels native close; only its exact process is disposable.
                if (identity is { } owned)
                {
                    Assert.True(OwnedProcessGuard.TryTerminate(owned, TimeSpan.Zero, TimeSpan.FromSeconds(5), out _));
                    if (retainedSession is not null)
                    {
                        await CloseSessionAsync(service, retainedSession, save: false);
                        sessions.TryRemove(retainedSession, out _);
                    }
                    Assert.DoesNotContain(owned, SessionManager.GetTrackedExcelProcesses());
                }
            }
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
