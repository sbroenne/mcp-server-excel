using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Workbook;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    [Theory]
    [InlineData("workbook.save-as")]
    [InlineData("session.save")]
    [InlineData("service.shutdown")]
    public async Task NativeModalPrompt_RefusesPersistenceBeforeQueuingComWork(string command)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "native-prompt.xlsx");
            var target = Path.Join(directory, "blocked-native-prompt.xlsx");
            var session = await CreateSessionAsync(service, path, show: true);
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, "Unsaved native prompt marker");
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(session));
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
                dismiss.TrySetResult(true);
                var rejected = Assert.Throws<COMException>(() => app.ScreenUpdating = false);
                Assert.Contains(rejected.HResult, new[]
                {
                    unchecked((int)0x80010001), unchecked((int)0x800AC472)
                });
                Assert.Equal(WorkbookRefreshState.DialogOpen, batch.GetExecutionPreflightState());
                ExcelBatch.WorkItemQueuedHookForTests = () => Interlocked.Increment(ref queued);
                if (command == "workbook.save-as")
                {
                    var busy = Assert.Throws<ExcelBusyException>(() => new WorkbookCommands().SaveAs(batch, target));
                    Assert.Contains("dialog", busy.Message, StringComparison.OrdinalIgnoreCase);
                }
                else
                {
                    var response = await service.ProcessAsync(new ServiceRequest { Command = command, SessionId = session });
                    Assert.False(response.Success);
                    Assert.Equal("Busy", response.ErrorCategory);
                }
                Assert.Equal(0, queued);
                Assert.False(File.Exists(target));
                Assert.Equal(path, batch.WorkbookPath);
                Assert.False(batch.HasTimedOutOperation);
            });
            ExcelBatch.WorkItemQueuedHookForTests = null;
            dismiss.TrySetResult(false);
            var dismissalFailure = await Record.ExceptionAsync(() => dismissal);
            var promptFailure = await Record.ExceptionAsync(async () => await prompt.WaitAsync(TimeSpan.FromSeconds(10)));
            var failures = new[] { failure, dismissalFailure, promptFailure }.OfType<Exception>().ToArray();
            if (failures.Length > 0)
                throw new AggregateException("Native persistence preflight or prompt cleanup failed.", failures);
            Assert.Equal("Unsaved native prompt marker", await ReadMarkerAsync(service, session));
            Assert.False(batch.Execute((context, _) => context.Book.Saved));
        });
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern nint FindWindow(string? className, string title);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(nint window, out uint processId);

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
    public async Task Shutdown_WithOwnedDialog_RetainsUnsavedWorkbookUntilSafeRetry()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "shutdown-dialog.xlsx");
            var session = await CreateSessionAsync(service, path);
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, "Retained through shutdown refusal");
            var batch = Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(session));
            var owner = batch.Execute((context, _) => new nint(context.App.Hwnd));
            using (var dialog = new OwnedTestDialog(owner))
            {
                var refusal = await service.ProcessAsync(new ServiceRequest { Command = "service.shutdown" });
                Assert.False(refusal.Success);
                Assert.Equal("Busy", refusal.ErrorCategory);
                Assert.Same(batch, service.SessionManager.GetSession(session));
                Assert.True((await service.ProcessAsync(new ServiceRequest { Command = "service.ping" })).Success);
                Assert.False(batch.HasTimedOutOperation);
            }
            Assert.Equal("Retained through shutdown refusal", await ReadMarkerAsync(service, session));
            Assert.False(batch.Execute((context, _) => context.Book.Saved));
            var retry = await service.ProcessAsync(new ServiceRequest { Command = "service.shutdown" });
            RequireSuccess(retry);
            sessions.TryRemove(session, out _);
            Assert.Equal(0, service.SessionManager.ActiveSessionCount);
            using var reopenedService = new ExcelMcpService();
            var reopened = await OpenSessionAsync(reopenedService, path);
            try
            {
                Assert.Equal("Retained through shutdown refusal", await ReadMarkerAsync(reopenedService, reopened));
            }
            finally
            {
                await CloseSessionAsync(reopenedService, reopened, save: false);
            }
        });
    }

    [Fact]
    public async Task ExcelOwnedDialog_ReportsPromptWithoutReadingContents_AndPreservesEdits()
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
                Assert.Contains("Check the Excel window", listed.GetProperty("blockingReason").GetString(), StringComparison.Ordinal);
                Assert.False(listed.GetProperty("canClose").GetBoolean());
                Assert.Equal(0, listed.GetProperty("activeOperations").GetInt32());
                Assert.DoesNotContain("private dialog contents", response.Result, StringComparison.Ordinal);
                Assert.DoesNotContain("sign-in required", response.Result, StringComparison.OrdinalIgnoreCase);
                var saveFailure = Assert.Throws<ExcelBusyException>(() => batch.Save());
                Assert.Contains("dialog", saveFailure.Message, StringComparison.Ordinal);

                var saveAs = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "workbook.save-as",
                    SessionId = session,
                    Args = JsonSerializer.Serialize(new { targetPath = Path.Join(directory, "blocked.xlsx") }, ServiceProtocol.JsonOptions)
                });
                Assert.False(saveAs.Success);
                Assert.Equal("Busy", saveAs.ErrorCategory);
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
                    Assert.Same(batch, service.SessionManager.GetSession(session));
                }
                Assert.False(batch.HasTimedOutOperation);
                Assert.True(timer.Elapsed < TimeSpan.FromSeconds(5), "Dialog inspection must not wait for Excel COM.");
            }

            var ready = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
            RequireSuccess(ready);
            using var readyJson = JsonDocument.Parse(ready.Result!);
            var readySession = Assert.Single(readyJson.RootElement.GetProperty("sessions").EnumerateArray());
            Assert.Equal("ready", readySession.GetProperty("excelState").GetString());
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
                window = CreateWindowEx(0x08000000, "STATIC", "private dialog contents", 0x90000000,
                    -10000, -10000, 1, 1, owner, 0, 0, 0);
                if (window == 0) throw new Win32Exception(Marshal.GetLastWin32Error());
                EnableWindow(owner, false);
                _opened.SetResult();
                _release.Wait();
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
    }
}
