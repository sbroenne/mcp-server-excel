using System.ComponentModel;
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
    [Theory]
    [InlineData("connection.discover-olap-schema")]
    [InlineData("connection.search-olap-members")]
    public async Task UnavailableOlapSession_ReturnsAutomaticDialogCheckWithoutChangingWorkbook(string command)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var session = await CreateSessionAsync(service, Path.Join(directory, "olap-dialog-check.xlsx"));
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, "Retained after discovery failure");
            var batch = Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(session));
            const string connectionName = "UnrefreshedOlap";
            batch.Execute((context, _) =>
            {
                Excel.Connections? connections = null;
                Excel.WorkbookConnection? connection = null;
                Excel.OLEDBConnection? oledb = null;
                try
                {
                    connections = context.Book.Connections;
                    connection = connections.Add2(connectionName, "Unrefreshed metadata",
                        "OLEDB;Provider=MSOLAP.8;Data Source=UnconfiguredRegressionSource;Initial Catalog=RegressionCube",
                        "Model", Excel.XlCmdType.xlCmdCube, false, false);
                    oledb = connection.OLEDBConnection;
                    Assert.True(oledb.OLAP);
                    Assert.False(oledb.Refreshing);
                    Assert.False(oledb.RefreshOnFileOpen);
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
                Args = JsonSerializer.Serialize(new
                {
                    connectionName,
                    hierarchyUniqueName = "[Test]",
                    levelUniqueName = "[Test].[Level]"
                }, ServiceProtocol.JsonOptions)
            });

            Assert.False(response.Success);
            Assert.Equal(session, response.SessionId);
            Assert.Equal(command, response.Command);
            Assert.Equal(nameof(InvalidOperationException), response.ExceptionType);
            Assert.Contains("active ADO session", response.ErrorMessage, StringComparison.Ordinal);
            Assert.Contains("No blocking Excel dialog was detected.", response.ErrorMessage, StringComparison.Ordinal);
            Assert.DoesNotContain("First, inspect this session", response.ErrorMessage, StringComparison.Ordinal);
            Assert.False(batch.HasTimedOutOperation);
            Assert.Equal("Retained after discovery failure", await ReadMarkerAsync(service, session));
            batch.Execute((context, _) =>
            {
                Excel.Connections? connections = null;
                Excel.WorkbookConnection? connection = null;
                Excel.OLEDBConnection? oledb = null;
                try
                {
                    connections = context.Book.Connections;
                    Assert.Equal(1, connections.Count);
                    connection = connections.Item(connectionName);
                    Assert.Equal(connectionName, connection.Name);
                    oledb = connection.OLEDBConnection;
                    Assert.Equal("Model", Assert.IsType<string>((object?)oledb.CommandText));
                    Assert.False(oledb.RefreshOnFileOpen);
                    Assert.False(oledb.Refreshing);
                    Assert.False(context.Book.Saved);
                }
                finally
                {
                    ComUtilities.Release(ref oledb);
                    ComUtilities.Release(ref connection);
                    ComUtilities.Release(ref connections);
                }
            });
        });
    }

    [Theory]
    [InlineData("workbook.save-as", false)]
    [InlineData("session.save", false)]
    [InlineData("service.shutdown", false)]
    [InlineData("session.close", false)]
    [InlineData("connection.get-refresh-status", false)]
    [InlineData("connection.cancel-refresh", false)]
    [InlineData("connection.get-account-settings", false)]
    [InlineData("connection.set-account-settings", false)]
    [InlineData("connection.clear-account-hint", false)]
    [InlineData("range.set-values", false)]
    [InlineData("range.set-formulas", true)]
    [InlineData("session.save", true)]
    [InlineData("workbook.save-as", true)]
    public async Task NativeModalPrompt_RefusesPersistenceBeforeQueuingComWork(string command, bool throughAttachment)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "native-prompt.xlsx");
            var target = Path.Join(directory, "blocked-native-prompt.xlsx");
            var session = await CreateSessionAsync(service, path, show: true);
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, "Unsaved native prompt marker");
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(session));
            var ownerSession = session;
            var marker = "Unsaved native prompt marker";
            if (throughAttachment)
            {
                marker = "Unsaved attached prompt marker";
                var name = batch.Execute((context, _) => AddReviewWorkbook(context.App, marker));
                session = await AttachReviewWorkbookAsync(service, await DiscoverReviewWorkbookAsync(service, name));
                sessions[session] = 0;
            }
            var selectedBatch = Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(session));
            var selectedPath = selectedBatch.WorkbookPath;
            var identity = Assert.Single(SessionManager.GetTrackedExcelProcesses(),
                item => item.ProcessId == batch.ExcelProcessId);
            var app = batch.Execute((context, _) => context.App);
            var title = $"ExcelMcp save preflight {Guid.NewGuid():N}";
            // Another apartment opens the real Excel prompt while the batch thread remains free.
            var prompt = Task.Run(() => app.GetSaveAsFilename(InitialFilename: target, Title: title));
            nint window = 0;
            var dismiss = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            var dismissal = Task.Run(async () =>
            {
                if (await Task.WhenAny(prompt, dismiss.Task) == prompt) return;
                if (await dismiss.Task) await Task.Delay(TimeSpan.FromSeconds(2));
                var timer = System.Diagnostics.Stopwatch.StartNew();
                while (!prompt.IsCompleted && timer.Elapsed < TimeSpan.FromSeconds(10))
                {
                    var ownedWindow = FindWindow(null, title);
                    if (ownedWindow != 0 && GetWindowThreadProcessId(ownedWindow, out var pid) != 0
                        && pid == identity.ProcessId && OwnedProcessGuard.MatchesLiveProcess(identity))
                    {
                        Assert.True(PostMessage(ownedWindow, 0x0010, 0, 0), "Could not dismiss the owned prompt.");
                        return;
                    }
                    await Task.Delay(50);
                }
                Assert.True(prompt.IsCompleted, "The owned native prompt could not be found for cleanup.");
            });
            var queued = 0;
            var failure = await Record.ExceptionAsync(async () =>
            {
                Assert.True(SpinWait.SpinUntil(() =>
                {
                    window = FindWindow(null, title);
                    return window != 0 && GetWindowThreadProcessId(window, out var pid) != 0
                        && pid == identity.ProcessId && IsWindowVisible(window) && IsWindowEnabled(window)
                        && batch.GetExecutionPreflightState() == WorkbookRefreshState.DialogOpen;
                }, TimeSpan.FromSeconds(10)), "The owned native Excel prompt did not open.");
                ExcelBatch.WorkItemQueuedHookForTests = () => Interlocked.Increment(ref queued);
                var inspected = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.notices",
                    SessionId = session
                });
                RequireSuccess(inspected);
                using var inspection = JsonDocument.Parse(inspected.Result!);
                var notices = inspection.RootElement.GetProperty("notices");
                Assert.False(notices.GetProperty("complete").GetBoolean());
                Assert.Equal(JsonValueKind.Null, notices.GetProperty("protectedView").ValueKind);
                Assert.Equal(JsonValueKind.Null, notices.GetProperty("readOnly").ValueKind);
                Assert.Equal(JsonValueKind.Null, notices.GetProperty("hasVbaProject").ValueKind);
                Assert.Equal(JsonValueKind.Null, notices.GetProperty("connectionCount").ValueKind);
                Assert.Equal("unknown", notices.GetProperty("vbaState").GetString());
                Assert.Equal("unknown", notices.GetProperty("connectionsState").GetString());
                Assert.StartsWith("Dialog text",
                    Assert.Single(notices.GetProperty("messages").EnumerateArray()).GetString());
                dismiss.TrySetResult(true);
                var rejected = Assert.Throws<COMException>(() => app.ScreenUpdating = false);
                Assert.Contains(rejected.HResult, new[]
                {
                    unchecked((int)0x80010001), unchecked((int)0x800AC472)
                });
                Assert.Equal(WorkbookRefreshState.DialogOpen, batch.GetExecutionPreflightState());
                if (command == "workbook.save-as")
                {
                    var busy = Assert.Throws<ExcelBusyException>(() => new WorkbookCommands().SaveAs(selectedBatch, target));
                    Assert.Contains("dialog", busy.Message, StringComparison.OrdinalIgnoreCase);
                    Assert.Contains(ExcelBusyException.DialogGuidance, busy.Message, StringComparison.Ordinal);
                    Assert.Contains("Dialog text", busy.Message, StringComparison.Ordinal);
                }
                else
                {
                    var args = command == "range.set-values"
                        ? """{"sheetName":"Sheet1","rangeAddress":"A1","values":[["must-not-write"]]}"""
                        : command == "range.set-formulas"
                        ? """{"sheetName":"Sheet1","rangeAddress":"A1","formulas":[["=1+2"]]}"""
                        : command == "connection.set-account-settings"
                        ? JsonSerializer.Serialize(new { connectionName = "DialogProbe", interactiveLogin = "Enabled" },
                            ServiceProtocol.JsonOptions)
                        : command.StartsWith("connection.", StringComparison.Ordinal)
                            ? JsonSerializer.Serialize(new { connectionName = "DialogProbe" }, ServiceProtocol.JsonOptions)
                            : null;
                    var response = await service.ProcessAsync(new ServiceRequest
                    {
                        Command = command,
                        SessionId = session,
                        Args = args
                    });
                    if (command == "service.shutdown")
                    {
                        Assert.True(response.Success, response.ErrorMessage);
                        Assert.Contains("dialog", response.PartialState, StringComparison.OrdinalIgnoreCase);
                        Assert.Contains("Dialog text", response.PartialState, StringComparison.Ordinal);
                    }
                    else
                    {
                        Assert.False(response.Success);
                        Assert.Equal("Busy", response.ErrorCategory);
                        Assert.Contains(ExcelBusyException.DialogGuidance, response.ErrorMessage, StringComparison.Ordinal);
                        Assert.Contains("Dialog text", response.ErrorMessage, StringComparison.Ordinal);
                    }
                }
                Assert.Equal(0, queued);
                Assert.False(File.Exists(target));
                Assert.Equal(selectedPath, selectedBatch.WorkbookPath);
                Assert.False(batch.HasTimedOutOperation);
            });
            ExcelBatch.WorkItemQueuedHookForTests = null;
            dismiss.TrySetResult(false);
            var dismissalFailure = await Record.ExceptionAsync(() => dismissal);
            var promptFailure = await Record.ExceptionAsync(async () => await prompt.WaitAsync(TimeSpan.FromSeconds(10)));
            var failures = new[] { failure, dismissalFailure, promptFailure }.OfType<Exception>().ToArray();
            if (failures.Length > 0)
                throw new AggregateException("Native persistence preflight or prompt cleanup failed.", failures);
            if (command != "service.shutdown")
                Assert.Equal(marker, await ReadMarkerAsync(service, session));
            Assert.False(selectedBatch.Execute((context, _) => context.Book.Saved));
            if (throughAttachment)
            {
                Assert.Equal("Unsaved native prompt marker", await ReadMarkerAsync(service, ownerSession));
                Assert.False(batch.Execute((context, _) => context.Book.Saved));
            }
        });
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern nint FindWindow(string? className, string title);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(nint window, out uint processId);

    private static readonly string[] DialogVisibilityActions = ["hide", "show", "hide"];
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowVisible(nint window);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowEnabled(nint window);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PostMessage(nint window, uint message, nint wParam, nint lParam);

    [Fact]
    public async Task PendingOperation_RevealsDialogWithoutPolling_AndSharesStandaloneShowHide()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var session = await CreateSessionAsync(service, Path.Join(directory, "pending-dialog.xlsx"));
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, "Still retained");
            var batch = Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(session));
            var owner = batch.Execute((context, _) => new nint(context.App.Hwnd));
            Assert.False(IsWindowVisible(owner));
            using var entered = new ManualResetEventSlim();
            using var release = new ManualResetEventSlim();
            using var dialog = new OwnedTestDialog(owner);
            var pending = Task.Run(() => batch.Execute((_, _) =>
            {
                entered.Set();
                Assert.True(release.Wait(TimeSpan.FromSeconds(15)));
                return 17;
            }));
            try
            {
                Assert.True(entered.Wait(TimeSpan.FromSeconds(5)));
                Assert.True(SpinWait.SpinUntil(() => IsWindowVisible(owner), TimeSpan.FromSeconds(3)),
                    "The operation's own wait must reveal Excel without a session-list request.");
                Assert.False(pending.IsCompleted);
                foreach (var action in DialogVisibilityActions)
                {
                    var response = await service.ProcessAsync(new ServiceRequest
                    {
                        Command = "window." + action,
                        SessionId = session
                    }).WaitAsync(TimeSpan.FromSeconds(3));
                    RequireSuccess(response);
                    using var json = JsonDocument.Parse(response.Result!);
                    Assert.Equal(action == "show", json.RootElement.GetProperty("isVisible").GetBoolean());
                    Assert.Equal(action == "show", IsWindowVisible(owner));
                }
                await Task.Delay(1100);
                Assert.False(IsWindowVisible(owner));
                Assert.False(pending.IsCompleted);
                Assert.False(batch.HasTimedOutOperation);
            }
            finally
            {
                release.Set();
                Assert.Equal(17, await pending.WaitAsync(TimeSpan.FromSeconds(5)));
            }
            dialog.Dispose();
            Assert.Equal("Still retained", await ReadMarkerAsync(service, session));
            Assert.False(batch.Execute((context, _) => context.Book.Saved));
            Assert.False(batch.Execute((context, _) => context.App.Visible));
            RequireSuccess(await service.ProcessAsync(new ServiceRequest { Command = "window.show", SessionId = session }));
            Assert.True(batch.Execute((context, _) => context.App.Visible));
            RequireSuccess(await service.ProcessAsync(new ServiceRequest { Command = "window.hide", SessionId = session }));
        });
    }

    [Fact]
    public async Task TimedOutOperation_WithOwnedDialog_AllowsOnlyNativeShowHide()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var session = service.SessionManager.CreateSessionForNewFile(
                Path.Join(directory, "timed-out-dialog.xlsx"),
                operationTimeout: TimeSpan.FromSeconds(15));
            sessions[session] = 0;
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(session));
            var owner = batch.Execute((context, _) => new nint(context.App.Hwnd));
            using var dialog = new OwnedTestDialog(owner);
            using var entered = new ManualResetEventSlim();
            using var release = new ManualResetEventSlim();
            var finished = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var pending = Task.Run(() => batch.Execute((_, _) =>
            {
                entered.Set();
                try
                {
                    Assert.True(release.Wait(TimeSpan.FromSeconds(30)));
                    return 17;
                }
                finally
                {
                    finished.TrySetResult();
                }
            }));

            try
            {
                Assert.True(entered.Wait(TimeSpan.FromSeconds(5)));
                Assert.Throws<TimeoutException>(() => pending.GetAwaiter().GetResult());
                Assert.True(batch.HasTimedOutOperation);
                Assert.True(IsWindowVisible(owner), "Timeout recovery must reveal Excel.");
                RequireSuccess(await service.ProcessAsync(new ServiceRequest { Command = "window.hide", SessionId = session })
                    .WaitAsync(TimeSpan.FromSeconds(3)));
                Assert.False(IsWindowVisible(owner));
            }
            finally
            {
                release.Set();
            }
            await finished.Task.WaitAsync(TimeSpan.FromSeconds(5));
            Assert.True(SpinWait.SpinUntil(() => batch.OperationStatus.State is "completed" or "failed", TimeSpan.FromSeconds(5)));
            Assert.Equal("completed", batch.OperationStatus.State);
            Assert.False(IsWindowVisible(owner), "Finishing the timed-out operation must not undo an explicit hide.");

            var rejectedWorkbookCommand = await service.ProcessAsync(new ServiceRequest
            {
                Command = "range.get-values",
                SessionId = session,
                Args = """{"sheetName":"Sheet1","rangeAddress":"A1"}"""
            });
            Assert.False(rejectedWorkbookCommand.Success);
            Assert.Equal("SessionUnavailable", rejectedWorkbookCommand.ErrorCategory);

            foreach (var action in DialogVisibilityActions)
            {
                var response = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "window." + action,
                    SessionId = session
                });
                RequireSuccess(response);
                using var json = JsonDocument.Parse(response.Result!);
                var expectedVisible = action == "show";
                Assert.Equal(expectedVisible, json.RootElement.GetProperty("isVisible").GetBoolean());
                Assert.Equal(expectedVisible, IsWindowVisible(owner));
            }
        });
    }

    [Fact]
    public async Task Shutdown_WithOwnedDialog_StopsAndLeavesUnsavedWorkbookAsIs()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "shutdown-dialog.xlsx");
            var session = await CreateSessionAsync(service, path);
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, "Left through shutdown");
            var batch = Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(session));
            var owner = batch.Execute((context, _) => new nint(context.App.Hwnd));
            using (var dialog = new OwnedTestDialog(owner))
            {
                var stopped = await service.ProcessAsync(new ServiceRequest { Command = "service.shutdown" });
                Assert.True(stopped.Success, stopped.ErrorMessage);
                Assert.Contains($"Excel process {batch.ExcelProcessId}", stopped.PartialState, StringComparison.Ordinal);
                Assert.Same(batch, service.SessionManager.GetSession(session));
                Assert.False(batch.HasTimedOutOperation);
                Assert.True(IsWindowVisible(owner));
            }
            Assert.False(batch.Execute((context, _) => context.Book.Saved));
            service.SessionManager.Dispose();
            sessions.TryRemove(session, out _);
            Assert.Equal(0, service.SessionManager.ActiveSessionCount);
            using var reopenedService = new ExcelMcpService();
            var reopened = await AttachHandoverAsync(reopenedService, path, batch.ExcelProcessId);
            try
            {
                Assert.Equal("Left through shutdown", await ReadMarkerAsync(reopenedService, reopened));
                Assert.False(reopenedService.SessionManager.GetSession(reopened)!.Execute((context, _) => context.Book.Saved));
            }
            finally
            {
                await CloseSessionAsync(reopenedService, reopened, save: false);
            }
        });
    }

    [Fact]
    public async Task ExcelOwnedDialog_ReportsReadableTextAutomatically_AndPreservesEdits()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "dialog-state.xlsx");
            var session = await CreateSessionAsync(service, path);
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, "Preserved after dialog");
            var batch = Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(session));
            var owner = batch.Execute((context, _) => new nint(context.App.Hwnd));
            using (var dialog = new OwnedTestDialog(owner))
            {
                var timer = System.Diagnostics.Stopwatch.StartNew();
                var response = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
                RequireSuccess(response);
                using var json = JsonDocument.Parse(response.Result!);
                var listed = Assert.Single(json.RootElement.GetProperty("sessions").EnumerateArray());
                Assert.Equal("dialogOpen", listed.GetProperty("excelState").GetString());
                Assert.Equal("dialogOpen", listed.GetProperty("blockingReason").GetString());
                Assert.Contains("Check the Excel window", listed.GetProperty("blockingMessage").GetString(), StringComparison.Ordinal);
                Assert.True(listed.GetProperty("isExcelVisible").GetBoolean(),
                    listed.GetProperty("blockingMessage").GetString());
                Assert.True(IsWindowVisible(owner));
                Assert.False(listed.GetProperty("canClose").GetBoolean());
                Assert.Equal(0, listed.GetProperty("activeOperations").GetInt32());
                Assert.Contains("Connection test message", listed.GetProperty("blockingMessage").GetString(), StringComparison.Ordinal);
                Assert.Contains("The database is unavailable.", listed.GetProperty("blockingMessage").GetString(), StringComparison.Ordinal);
                Assert.DoesNotContain("test-entered-password", response.Result, StringComparison.Ordinal);
                Assert.DoesNotContain("sign-in required", response.Result, StringComparison.OrdinalIgnoreCase);
                var saveFailure = Assert.Throws<ExcelBusyException>(() => batch.Save());
                Assert.Contains("dialog", saveFailure.Message, StringComparison.Ordinal);
                Assert.Contains("Connection test message", saveFailure.Message, StringComparison.Ordinal);
                Assert.Contains("The database is unavailable.", saveFailure.Message, StringComparison.Ordinal);
                Assert.DoesNotContain("test-entered-password", saveFailure.Message, StringComparison.Ordinal);

                var saveAs = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "workbook.save-as",
                    SessionId = session,
                    Args = JsonSerializer.Serialize(new { targetPath = Path.Join(directory, "blocked.xlsx") }, ServiceProtocol.JsonOptions)
                });
                Assert.False(saveAs.Success);
                Assert.Equal("Busy", saveAs.ErrorCategory);
                Assert.Contains("Connection test message", saveAs.ErrorMessage, StringComparison.Ordinal);
                Assert.Contains("The database is unavailable.", saveAs.ErrorMessage, StringComparison.Ordinal);
                Assert.DoesNotContain("test-entered-password", saveAs.ErrorMessage, StringComparison.Ordinal);
                Assert.False(File.Exists(Path.Join(directory, "blocked.xlsx")));
                foreach (var save in CloseSaveModes)
                {
                    var close = await service.ProcessAsync(new ServiceRequest
                    {
                        Command = "session.close",
                        SessionId = session,
                        Args = JsonSerializer.Serialize(new { save }, ServiceProtocol.JsonOptions)
                    });
                    Assert.False(close.Success);
                    Assert.Equal("Busy", close.ErrorCategory);
                    Assert.Contains("dialog", close.ErrorMessage, StringComparison.Ordinal);
                    Assert.Contains("Connection test message", close.ErrorMessage, StringComparison.Ordinal);
                    Assert.Contains("The database is unavailable.", close.ErrorMessage, StringComparison.Ordinal);
                    Assert.DoesNotContain("test-entered-password", close.ErrorMessage, StringComparison.Ordinal);
                    Assert.Same(batch, service.SessionManager.GetSession(session));
                }
                Assert.False(batch.HasTimedOutOperation);
                Assert.True(timer.Elapsed < TimeSpan.FromSeconds(5), "Dialog inspection must not wait for Excel COM.");
            }

            JsonElement readySession;
            var readyTimer = System.Diagnostics.Stopwatch.StartNew();
            do
            {
                var ready = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
                RequireSuccess(ready);
                using var readyJson = JsonDocument.Parse(ready.Result!);
                readySession = Assert.Single(readyJson.RootElement.GetProperty("sessions").EnumerateArray()).Clone();
                if (readySession.GetProperty("excelState").GetString() == "ready") break;
                await Task.Delay(50);
            } while (readyTimer.Elapsed < TimeSpan.FromSeconds(10));
            Assert.Equal("ready", readySession.GetProperty("excelState").GetString());
            Assert.False(readySession.TryGetProperty("blockingMessage", out _));
            Assert.True(readySession.GetProperty("canClose").GetBoolean());
            Assert.Equal("Preserved after dialog", await ReadMarkerAsync(service, session));
            Assert.False(batch.Execute((context, _) => context.Book.Saved));
            await CloseSessionAsync(service, session, save: true);
            sessions.TryRemove(session, out _);
            var reopened = await OpenSessionAsync(service, path);
            sessions[reopened] = 0;
            Assert.Equal("Preserved after dialog", await ReadMarkerAsync(service, reopened));
        });
    }

    private sealed class OwnedTestDialog : IDisposable
    {
        private int _disposed;
        private readonly ManualResetEventSlim _release = new();
        private readonly Thread _thread;
        private readonly TaskCompletionSource _opened = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource _finished = new(TaskCreationOptions.RunContinuationsAsynchronously);

        internal OwnedTestDialog(nint owner)
        {
            _thread = new Thread(() => Run(owner)) { IsBackground = true };
            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
            try
            {
                _opened.Task.WaitAsync(TimeSpan.FromSeconds(10)).GetAwaiter().GetResult();
            }
            catch (Exception primary)
            {
                try { Dispose(); }
                catch (Exception cleanup) { throw new AggregateException(primary, cleanup); }
                throw;
            }
        }

        private void Run(nint owner)
        {
            nint window = 0;
            var failures = new List<Exception>();
            try
            {
                window = CreateWindowEx(0x08000000, "STATIC", "Connection test message", 0x90000000,
                    -10000, -10000, 1, 1, owner, 0, 0, 0);
                if (window == 0) throw new Win32Exception(Marshal.GetLastWin32Error());
                if (CreateWindowEx(0, "STATIC", "The database is unavailable.", 0x50000000,
                    0, 0, 1, 1, window, 0, 0, 0) == 0 ||
                    CreateWindowEx(0, "EDIT", "test-entered-password", 0x50000000,
                    0, 0, 1, 1, window, 0, 0, 0) == 0)
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                EnableWindow(owner, false);
                _opened.SetResult();
                // Owned windows share input processing with Excel; keep the dialog responsive.
                while (!_release.Wait(10))
                {
                    while (PeekMessage(out var message, 0, 0, 0, 1))
                        _ = DispatchMessage(ref message);
                }
            }
            catch (Exception ex)
            {
                failures.Add(ex);
                _opened.TrySetException(ex);
            }
            finally
            {
                EnableWindow(owner, true);
                if (window != 0 && !DestroyWindow(window))
                    failures.Add(new Win32Exception("Failed to destroy the owned test dialog."));
            }
            if (failures.Count > 0) _finished.SetException(failures);
            else _finished.SetResult();
        }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
            _release.Set();
            if (!_thread.Join(TimeSpan.FromSeconds(10)))
                throw new TimeoutException("Owned dialog thread did not finish cleanup.");
            _release.Dispose();
            _finished.Task.GetAwaiter().GetResult();
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern nint CreateWindowEx(uint extendedStyle, string className, string title,
            uint style, int x, int y, int width, int height, nint owner, nint menu, nint instance, nint parameter);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnableWindow(nint window, [MarshalAs(UnmanagedType.Bool)] bool enabled);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool DestroyWindow(nint window);

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeMessage
        {
            public nint Window;
            public uint Message;
            public nint WParam;
            public nint LParam;
            public uint Time;
            public int X;
            public int Y;
            public uint Private;
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool PeekMessage(out NativeMessage message, nint window, uint min, uint max, uint remove);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern nint DispatchMessage(ref NativeMessage message);
    }
}
