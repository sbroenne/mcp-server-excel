using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
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
