using System.Runtime.ExceptionServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Tests.Unit;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Service;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Sequential")]
[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "ServiceDaemon")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Medium")]
public sealed class CliServiceTrayExitTests
{
    [Fact]
    public async Task Exit_AfterWorkbookWasClosedInExcel_DoesNotReportFailure()
    {
        await RunWithCleanupAsync(async (service, directory, _) =>
        {
            var sessionId = await OpenAsync(service, "session.create", Path.Join(directory, "already-closed.xlsx"));
            var batch = GetBatch(service, sessionId);
            batch.Execute((context, _) => context.Book.Close(SaveChanges: false));
            Assert.True(batch.IsExcelProcessAlive());
            var shutdowns = 0;
            var notifications = new List<(string Title, string Message, ToolTipIcon Icon)>();

            await CliServiceTrayTestHost.RunAsync(
                () => CreateTray(service, DialogResult.No, [], notifications, () => shutdowns++),
                tray => tray.ExitServiceAsync());

            Assert.Equal(1, shutdowns);
            Assert.Empty(notifications);
            Assert.Empty(service.SessionManager.ActiveSessionIds);
            Assert.False(batch.IsExcelProcessAlive());
        });
    }

    [Fact]
    public async Task Exit_AfterWorkbookWasClosedInExcel_ShowsExcelLeftOpenWithAnotherWorkbook()
    {
        await RunWithCleanupAsync(async (service, directory, cleanup) =>
        {
            var sessionId = await OpenAsync(service, "session.create", Path.Join(directory, "closed-by-user.xlsx"));
            var batch = GetBatch(service, sessionId);
            var neighborName = batch.Execute((context, _) =>
            {
                Excel.Workbooks? books = null;
                Excel.Workbook? book = null;
                try
                {
                    books = context.App.Workbooks;
                    book = books.Add();
                    SetMarker(book, "keep neighboring workbook");
                    return book.Name;
                }
                finally
                {
                    ComUtilities.Release(ref book);
                    ComUtilities.Release(ref books);
                }
            });
            var observer = new SessionManager();
            cleanup.Add(observer.Dispose);
            cleanup.Add(() => CloseRetainedWorkbook(observer, batch, neighborName));
            batch.Execute((context, _) => context.Book.Close(SaveChanges: false));
            using var excel = System.Diagnostics.Process.GetProcessById(batch.ExcelProcessId!.Value);
            Assert.Equal(IntPtr.Zero, excel.MainWindowHandle);
            var shutdowns = 0;
            var notifications = new List<(string Title, string Message, ToolTipIcon Icon)>();

            await CliServiceTrayTestHost.RunAsync(
                () => CreateTray(service, DialogResult.No, [], notifications, () => shutdowns++),
                tray => tray.ExitServiceAsync());

            Assert.Equal(1, shutdowns);
            Assert.DoesNotContain(notifications, notification => notification.Icon == ToolTipIcon.Error);
            Assert.True(batch.IsExcelProcessAlive());
            Assert.True(SpinWait.SpinUntil(() =>
            {
                excel.Refresh();
                return excel.MainWindowHandle != IntPtr.Zero;
            }, TimeSpan.FromSeconds(5)), "Excel left open with another workbook must be visible.");
            var retained = AttachRetainedWorkbook(observer, batch, neighborName);
            AssertWorkbook(retained.Batch, "keep neighboring workbook", saved: false, path: string.Empty);
        });
    }

    [Fact]
    public async Task Exit_WhileSessionCloseIsInProgress_AsksWithoutWaitingForIt()
    {
        await RunWithCleanupAsync(async (service, directory, cleanup) =>
        {
            var sessionId = await OpenAsync(service, "session.create", Path.Join(directory, "closing.xlsx"));
            var batch = GetBatch(service, sessionId);
            batch.Execute((context, _) => context.Book.Close(SaveChanges: false));
            var disposeLock = (Lock)batch.GetType()
                .GetField("_disposeLock", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
                .GetValue(batch)!;
            var lockHeld = new ManualResetEventSlim();
            var releaseLock = new ManualResetEventSlim();
            cleanup.Add(lockHeld.Dispose);
            cleanup.Add(releaseLock.Dispose);
            var holder = new Thread(() =>
            {
                disposeLock.Enter();
                try
                {
                    lockHeld.Set();
                    releaseLock.Wait(TimeSpan.FromSeconds(30));
                }
                finally
                {
                    disposeLock.Exit();
                }
            })
            { IsBackground = true };
            cleanup.Add(() =>
            {
                releaseLock.Set();
                Assert.True(holder.Join(TimeSpan.FromSeconds(10)), "Dispose-lock owner did not exit.");
            });
            holder.Start();
            Assert.True(lockHeld.Wait(TimeSpan.FromSeconds(10)));
            var prompts = new List<string>();
            var shutdowns = 0;
            var notifications = new List<(string Title, string Message, ToolTipIcon Icon)>();
            var elapsed = System.Diagnostics.Stopwatch.StartNew();

            await CliServiceTrayTestHost.RunAsync(
                () => CreateTray(service, DialogResult.Cancel, prompts, notifications, () => shutdowns++),
                async tray =>
                {
                    await tray.ExitServiceAsync();
                    Assert.True(elapsed.Elapsed < TimeSpan.FromSeconds(3),
                        $"Exit waited {elapsed.Elapsed} for the close in progress.");
                });

            Assert.Equal(0, shutdowns);
            Assert.Empty(notifications);
            Assert.Contains("There are 1 sessions opened by ExcelMCP.", Assert.Single(prompts), StringComparison.Ordinal);
            Assert.Contains(sessionId, service.SessionManager.ActiveSessionIds);
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Exit_LeavesAttachedWorkbookOpenWithoutSaving(bool save)
    {
        await RunWithCleanupAsync(async (service, directory, cleanup) =>
        {
            var owner = new SessionManager();
            cleanup.Add(owner.Dispose);
            var ownerId = owner.CreateSessionForNewFile(Path.Join(directory, "owner.xlsx"));
            var ownerBatch = Assert.IsAssignableFrom<IExcelBatch>(owner.GetSession(ownerId));
            var name = ownerBatch.Execute((context, _) =>
            {
                Excel.Workbooks? books = null;
                Excel.Workbook? book = null;
                try
                {
                    books = context.App.Workbooks;
                    book = books.Add();
                    SetMarker(book, "keep attached edits");
                    return book.Name;
                }
                finally
                {
                    ComUtilities.Release(ref book);
                    ComUtilities.Release(ref books);
                }
            });
            cleanup.Add(() => CloseNeighbor(ownerBatch, name));
            var discovery = await service.SessionManager.DiscoverOpenWorkbooksAsync();
            var selection = Assert.Single(discovery.Workbooks, book =>
                book.Name == name && book.ProcessId == ownerBatch.ExcelProcessId);
            var attached = service.SessionManager.AttachWorkbook(selection.WorkbookId);
            AssertWorkbook(GetBatch(service, attached.SessionId), "keep attached edits", saved: false, path: string.Empty);
            var ownedPath = Path.Join(directory, "tray-owned.xlsx");
            var ownedId = await OpenAsync(service, "session.create", ownedPath);
            await WriteMarkerAsync(service, ownedId, "owned saved marker");
            GetBatch(service, ownedId).Save();
            var ownedMarker = save ? "owned exit save marker" : "owned saved marker";
            if (save)
                await WriteMarkerAsync(service, ownedId, ownedMarker);
            AssertWorkbook(GetBatch(service, ownedId), ownedMarker, saved: !save, directory);
            var prompts = new List<string>();
            var shutdowns = 0;
            var notifications = new List<(string Title, string Message, ToolTipIcon Icon)>();
            var closes = new List<(string SessionId, bool Save, ServiceResponse Response)>();

            await CliServiceTrayTestHost.RunAsync(
                () => CreateTray(service, save ? DialogResult.Yes : DialogResult.No, prompts, notifications,
                    () => shutdowns++, closes),
                tray => tray.ExitServiceAsync());

            Assert.Equal(1, shutdowns);
            Assert.Empty(notifications);
            Assert.Contains("There are 1 sessions opened by ExcelMCP.", Assert.Single(prompts), StringComparison.Ordinal);
            if (save)
            {
                var close = Assert.Single(closes);
                Assert.Equal(ownedId, close.SessionId);
                Assert.True(close.Save);
                RequireSuccess(close.Response);
            }
            else
            {
                Assert.Empty(closes);
            }
            ownerBatch.Execute((context, _) =>
            {
                Excel.Workbooks? books = null;
                Excel.Workbook? book = null;
                try
                {
                    books = context.App.Workbooks;
                    Assert.Equal(2, books.Count);
                    book = books[name];
                    Assert.False(book.Saved);
                    Assert.Equal(string.Empty, book.Path);
                    Assert.Equal("keep attached edits", ReadMarker(book));
                }
                finally
                {
                    ComUtilities.Release(ref book);
                    ComUtilities.Release(ref books);
                }
            });
            var verifier = new ExcelMcpService();
            cleanup.Add(verifier.Dispose);
            var reopened = await OpenAsync(verifier, "session.open", ownedPath);
            AssertWorkbook(GetBatch(verifier, reopened), ownedMarker, saved: true, directory);
        });
    }

    [Fact]
    public async Task Exit_No_LeavesOwnedUnsavedEditsOpenAndSavedContentsUnchanged()
    {
        await RunWithCleanupAsync(async (service, directory, cleanup) =>
        {
            var path = Path.Join(directory, "unsaved.xlsx");
            var sessionId = await OpenAsync(service, "session.create", path);
            await WriteMarkerAsync(service, sessionId, "saved baseline");
            var batch = GetBatch(service, sessionId);
            batch.Save();
            await WriteMarkerAsync(service, sessionId, "unsaved owned marker");
            AssertWorkbook(batch, "unsaved owned marker", saved: false, directory);
            var observer = new SessionManager();
            cleanup.Add(observer.Dispose);
            cleanup.Add(() => CloseRetainedWorkbook(observer, batch, Path.GetFileName(path)));
            var prompts = new List<string>();
            var notifications = new List<(string Title, string Message, ToolTipIcon Icon)>();
            var shutdowns = 0;
            var closes = new List<(string SessionId, bool Save, ServiceResponse Response)>();

            await CliServiceTrayTestHost.RunAsync(
                () => CreateTray(service, DialogResult.No, prompts, notifications, () => shutdowns++, closes),
                tray => tray.ExitServiceAsync());

            Assert.Equal(1, shutdowns);
            Assert.Single(prompts);
            Assert.Empty(closes);
            Assert.DoesNotContain(notifications, notification => notification.Icon == ToolTipIcon.Error);
            var retained = AttachRetainedWorkbook(observer, batch, Path.GetFileName(path));
            AssertWorkbook(retained.Batch, "unsaved owned marker", saved: false, directory);
            Assert.True(retained.Batch.Execute((context, _) => context.App.Visible));
            observer.CloseSession(retained.SessionId, save: false);
            var verifier = new ExcelMcpService();
            cleanup.Add(verifier.Dispose);
            var reopened = await OpenAsync(verifier, "session.open", path);
            AssertWorkbook(GetBatch(verifier, reopened), "saved baseline", saved: true, directory);
        });
    }

    [Fact]
    public async Task Exit_Cancel_PreservesUnsavedWorkbookAndAcceptsFurtherWork()
    {
        await RunWithCleanupAsync(async (service, directory, _) =>
        {
            var sessionId = await OpenAsync(service, "session.create", Path.Join(directory, "cancel.xlsx"));
            await WriteMarkerAsync(service, sessionId, "saved baseline");
            var batch = GetBatch(service, sessionId);
            batch.Save();
            await WriteMarkerAsync(service, sessionId, "cancel keeps edits");
            var prompts = new List<string>();
            var notifications = new List<(string Title, string Message, ToolTipIcon Icon)>();
            var shutdowns = 0;
            var closes = new List<(string SessionId, bool Save, ServiceResponse Response)>();

            await CliServiceTrayTestHost.RunAsync(
                () => CreateTray(service, DialogResult.Cancel, prompts, notifications, () => shutdowns++, closes),
                async tray =>
                {
                    await tray.ExitServiceAsync();
                    await tray.ExitServiceAsync();
                });

            Assert.Equal(2, prompts.Count);
            Assert.Equal(0, shutdowns);
            Assert.Empty(closes);
            Assert.Empty(notifications);
            Assert.Same(batch, service.SessionManager.GetSession(sessionId));
            AssertWorkbook(batch, "cancel keeps edits", saved: false, directory);
            RequireSuccess(await service.ProcessAsync(new ServiceRequest { Command = "session.list" }));
            RequireSuccess(await CloseAsync(service, sessionId, save: false));
        });
    }

    [Fact]
    public async Task Exit_WhenExcelCancelsRequestedSave_ReportsFailureAndKeepsServiceAndEdits()
    {
        await RunWithCleanupAsync(async (service, directory, cleanup) =>
        {
            var path = Path.Join(directory, "cancelled-save.xlsx");
            var sessionId = await OpenAsync(service, "session.create", path);
            await WriteMarkerAsync(service, sessionId, "saved baseline");
            RequireSuccess(await CloseAsync(service, sessionId, save: true));
            sessionId = await OpenAsync(service, "session.open", path);
            var batch = GetBatch(service, sessionId);
            await WriteMarkerAsync(service, sessionId, "failed save retains marker");
            AssertWorkbook(batch, "failed save retains marker", saved: false, directory);
            Assert.Equal(path, batch.Execute((context, _) => context.Book.FullName), ignoreCase: true);
            var eventCount = 0;
            Excel.AppEvents_WorkbookBeforeSaveEventHandler cancelSave =
                (Excel.Workbook _, bool _, ref bool cancel) =>
                {
                    Interlocked.Increment(ref eventCount);
                    cancel = true;
                };
            batch.Execute((context, _) => context.App.WorkbookBeforeSave += cancelSave);
            var subscribed = true;
            cleanup.Add(() =>
            {
                if (subscribed)
                    batch.Execute((context, _) => context.App.WorkbookBeforeSave -= cancelSave);
            });
            var prompts = new List<string>();
            var notifications = new List<(string Title, string Message, ToolTipIcon Icon)>();
            var closes = new List<(string SessionId, bool Save, ServiceResponse Response)>();
            var shutdowns = 0;

            await CliServiceTrayTestHost.RunAsync(
                () => CreateTray(service, DialogResult.Yes, prompts, notifications, () => shutdowns++, closes),
                async tray =>
                {
                    await tray.ExitServiceAsync();
                    Assert.Equal(0, shutdowns);
                    var close = Assert.Single(closes);
                    Assert.Equal(sessionId, close.SessionId);
                    Assert.True(close.Save);
                    Assert.False(close.Response.Success);
                    Assert.Contains("not saved", close.Response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
                    var notification = Assert.Single(notifications);
                    Assert.Equal("Error", notification.Title);
                    Assert.Equal(ToolTipIcon.Error, notification.Icon);
                    Assert.Contains(close.Response.ErrorMessage!, notification.Message, StringComparison.Ordinal);
                    Assert.True(Volatile.Read(ref eventCount) > 0, "Excel never reached WorkbookBeforeSave.");
                    Assert.Same(batch, service.SessionManager.GetSession(sessionId));
                    AssertWorkbook(batch, "failed save retains marker", saved: false, directory);
                    RequireSuccess(await service.ProcessAsync(new ServiceRequest { Command = "session.list" }));
                    await WriteMarkerAsync(service, sessionId, "service still accepts edits");
                    AssertWorkbook(batch, "service still accepts edits", saved: false, directory);
                    await tray.ExitServiceAsync();
                });

            Assert.Equal(2, prompts.Count);
            Assert.Equal(2, closes.Count);
            Assert.All(closes, close => Assert.False(close.Response.Success));
            Assert.Equal(2, notifications.Count);
            Assert.Equal(0, shutdowns);
            AssertWorkbook(batch, "service still accepts edits", saved: false, directory);
            batch.Execute((context, _) => context.App.WorkbookBeforeSave -= cancelSave);
            subscribed = false;
            RequireSuccess(await CloseAsync(service, sessionId, save: false));
            var reopened = await OpenAsync(service, "session.open", path);
            AssertWorkbook(GetBatch(service, reopened), "saved baseline", saved: true, directory);
        });
    }

    private static CliServiceTray CreateTray(
        ExcelMcpService service,
        DialogResult choice,
        List<string> prompts,
        List<(string Title, string Message, ToolTipIcon Icon)> notifications,
        Action shutdown,
        List<(string SessionId, bool Save, ServiceResponse Response)>? closes = null) =>
        new(() => service.SessionManager.GetActiveSessions(statusOnly: true),
            async (sessionId, save) =>
            {
                var response = await CloseAsync(service, sessionId, save);
                closes?.Add((sessionId, save, response));
                return response;
            },
            () =>
            {
                shutdown();
                return Task.FromResult(service.RequestShutdown());
            },
            (message, title, buttons, icon) =>
            {
                Assert.Equal("Exit ExcelMCP CLI", title);
                Assert.Equal(MessageBoxButtons.YesNoCancel, buttons);
                Assert.Equal(MessageBoxIcon.Question, icon);
                prompts.Add(message);
                return choice;
            },
            (title, message, icon) => notifications.Add((title, message, icon)));

    private static async Task RunWithCleanupAsync(
        Func<ExcelMcpService, string, List<Action>, Task> test)
    {
        var directory = Path.Join(Path.GetTempPath(), $"CliServiceTrayExitTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        var service = new ExcelMcpService();
        var cleanup = new List<Action>
        {
            () => Directory.Delete(directory, recursive: true),
            service.Dispose,
            () => CloseOwnedTestSessions(service)
        };
        Exception? failure = null;
        try
        {
            await test(service, directory, cleanup);
        }
        catch (Exception ex)
        {
            failure = ex;
        }
        finally
        {
            for (var index = cleanup.Count - 1; index >= 0; index--)
            {
                try
                {
                    cleanup[index]();
                }
                catch (Exception ex)
                {
                    failure = failure is null ? ex : new AggregateException(failure, ex);
                }
            }
        }
        if (failure is not null)
            ExceptionDispatchInfo.Capture(failure).Throw();
    }

    private static void CloseOwnedTestSessions(ExcelMcpService service)
    {
        if (!service.SessionManager.ActiveSessionIds.Any())
            return;
        var sessions = service.SessionManager.GetActiveSessions(statusOnly: true).Where(session => !session.IsAttached);
        Exception? failure = null;
        foreach (var session in sessions)
        {
            try
            {
                service.SessionManager.CloseSession(session.SessionId, save: false);
            }
            catch (Exception ex)
            {
                failure = failure is null ? ex : new AggregateException(failure, ex);
            }
        }
        if (failure is not null)
            ExceptionDispatchInfo.Capture(failure).Throw();
    }

    private static async Task<string> OpenAsync(ExcelMcpService service, string command, string path)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = command,
            Args = JsonSerializer.Serialize(new { filePath = path })
        });
        RequireSuccess(response);
        using var result = JsonDocument.Parse(response.Result!);
        var sessionId = result.RootElement.GetProperty("sessionId").GetString();
        Assert.False(string.IsNullOrWhiteSpace(sessionId));
        return sessionId!;
    }

    private static Task<ServiceResponse> CloseAsync(ExcelMcpService service, string sessionId, bool save) =>
        service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new { save })
        });

    private static async Task WriteMarkerAsync(ExcelMcpService service, string sessionId, string marker)
    {
        RequireSuccess(await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                sheetName = "Sheet1",
                rangeAddress = "A1",
                overwritePolicy = "allow",
                values = new object?[][] { [marker] }
            })
        }));
    }

    private static IExcelBatch GetBatch(ExcelMcpService service, string sessionId) =>
        Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(sessionId));

    private static void RequireSuccess(ServiceResponse response)
    {
        Assert.True(response.Success, response.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(response.ErrorMessage), response.ErrorMessage);
    }

    private static void AssertWorkbook(IExcelBatch batch, string marker, bool saved, string path) =>
        batch.Execute((context, _) =>
        {
            Assert.Equal(marker, ReadMarker(context.Book));
            Assert.Equal(saved, context.Book.Saved);
            Assert.Equal(path, context.Book.Path, ignoreCase: true);
        });

    private static string? ReadMarker(Excel.Workbook workbook)
    {
        Excel.Sheets? sheets = null;
        Excel.Worksheet? sheet = null;
        Excel.Range? cell = null;
        try
        {
            sheets = workbook.Worksheets;
            sheet = (Excel.Worksheet)sheets[1];
            cell = sheet.Range["A1"];
            return cell.Value2 as string;
        }
        finally
        {
            ComUtilities.Release(ref cell);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref sheets);
        }
    }

    private static void SetMarker(Excel.Workbook workbook, string marker)
    {
        Excel.Sheets? sheets = null;
        Excel.Worksheet? sheet = null;
        Excel.Range? cell = null;
        try
        {
            sheets = workbook.Worksheets;
            sheet = (Excel.Worksheet)sheets[1];
            cell = sheet.Range["A1"];
            cell.Value2 = marker;
        }
        finally
        {
            ComUtilities.Release(ref cell);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref sheets);
        }
    }

    private static (string SessionId, IExcelBatch Batch) AttachRetainedWorkbook(
        SessionManager observer, IExcelBatch original, string name)
    {
        var existing = observer.GetActiveSessions(statusOnly: true)
            .SingleOrDefault(session => session.WorkbookName == name);
        if (existing is not null)
            return (existing.SessionId, Assert.IsAssignableFrom<IExcelBatch>(observer.GetSession(existing.SessionId)));
        var discovery = observer.DiscoverOpenWorkbooksAsync().GetAwaiter().GetResult();
        var selection = Assert.Single(discovery.Workbooks, book =>
            book.Name == name && book.ProcessId == original.ExcelProcessId);
        var attached = observer.AttachWorkbook(selection.WorkbookId);
        return (attached.SessionId, Assert.IsAssignableFrom<IExcelBatch>(observer.GetSession(attached.SessionId)));
    }

    private static void CloseRetainedWorkbook(SessionManager observer, IExcelBatch original, string name)
    {
        if (!original.IsExcelProcessAlive())
            return;
        var existing = observer.GetActiveSessions(statusOnly: true)
            .SingleOrDefault(session => session.WorkbookName == name);
        if (existing is not null)
        {
            observer.CloseSession(existing.SessionId, save: false);
            return;
        }
        var discovery = observer.DiscoverOpenWorkbooksAsync().GetAwaiter().GetResult();
        var selection = discovery.Workbooks.SingleOrDefault(book =>
            book.Name == name && book.ProcessId == original.ExcelProcessId);
        if (selection is not null)
        {
            var attached = observer.AttachWorkbook(selection.WorkbookId);
            observer.CloseSession(attached.SessionId, save: false);
        }
    }

    private static void CloseNeighbor(IExcelBatch owner, string name) =>
        owner.Execute((context, _) =>
        {
            Excel.Workbooks? books = null;
            Excel.Workbook? book = null;
            try
            {
                books = context.App.Workbooks;
                for (var index = 1; index <= books.Count; index++)
                {
                    book = books[index];
                    if (book.Name == name)
                    {
                        book.Close(SaveChanges: false);
                        return;
                    }
                    ComUtilities.Release(ref book);
                }
            }
            finally
            {
                ComUtilities.Release(ref book);
                ComUtilities.Release(ref books);
            }
        });
}
