using System.Runtime.ExceptionServices;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
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
    public void Exit_AfterWorkbookWasClosedInExcel_DoesNotReportFailure()
    {
        var directory = Path.Join(Path.GetTempPath(), $"CliServiceTrayExitTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        var path = Path.Join(directory, "already-closed.xlsx");
        var manager = new SessionManager();
        IExcelBatch? batch = null;
        Exception? failure = null;
        try
        {
            var sessionId = manager.CreateSessionForNewFile(path);
            batch = Assert.IsAssignableFrom<IExcelBatch>(manager.GetSession(sessionId));
            batch.Execute((context, _) => context.Book.Close(SaveChanges: false));
            Assert.True(batch.IsExcelProcessAlive());

            var errors = new List<string>();
            var shutdown = false;
            CliServiceTray.ExitService(manager, () =>
            {
                manager.Dispose();
                shutdown = true;
                return null;
            }, (message, _, _, _) =>
            {
                errors.Add(message);
                return DialogResult.Cancel;
            }, (_, message, icon) =>
            {
                if (icon != ToolTipIcon.Info)
                    errors.Add(message);
            });

            Assert.True(shutdown, string.Join(Environment.NewLine, errors));
            Assert.Empty(errors);
            Assert.Empty(manager.ActiveSessionIds);
            Assert.False(batch.IsExcelProcessAlive());
        }
        catch (Exception ex)
        {
            failure = ex;
        }
        finally
        {
            Capture(manager.Dispose);
            Capture(() => Directory.Delete(directory, recursive: true));
        }
        if (failure is not null)
            ExceptionDispatchInfo.Capture(failure).Throw();

        void Capture(Action action)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                failure = failure is null ? ex : new AggregateException(failure, ex);
            }
        }
    }

    [Fact]
    public void Exit_WhileSessionCloseIsInProgress_AsksWithoutWaitingForIt()
    {
        var directory = Path.Join(Path.GetTempPath(), $"CliServiceTrayExitTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        var manager = new SessionManager();
        using var lockHeld = new ManualResetEventSlim();
        using var releaseLock = new ManualResetEventSlim();
        Thread? holder = null;
        Exception? failure = null;
        try
        {
            var sessionId = manager.CreateSessionForNewFile(Path.Join(directory, "closing.xlsx"));
            var batch = Assert.IsAssignableFrom<IExcelBatch>(manager.GetSession(sessionId));
            batch.Execute((context, _) => context.Book.Close(SaveChanges: false));
            var disposeLock = (Lock)batch.GetType()
                .GetField("_disposeLock", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
                .GetValue(batch)!;
            holder = new Thread(() =>
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
            });
            holder.Start();
            Assert.True(lockHeld.Wait(TimeSpan.FromSeconds(10)));

            var prompts = new List<string>();
            var elapsed = System.Diagnostics.Stopwatch.StartNew();
            var report = CliServiceTray.ExitService(manager, () => throw new InvalidOperationException("Shutdown must not start."),
                (message, _, _, _) =>
                {
                    prompts.Add(message);
                    return DialogResult.Cancel;
                }, (_, _, _) => { });
            elapsed.Stop();

            Assert.True(elapsed.Elapsed < TimeSpan.FromSeconds(3), $"Exit waited {elapsed.Elapsed} for the close in progress.");
            Assert.Null(report);
            Assert.Contains("1 active session(s)", Assert.Single(prompts), StringComparison.Ordinal);
            Assert.Contains(sessionId, manager.ActiveSessionIds);
        }
        catch (Exception ex)
        {
            failure = ex;
        }
        finally
        {
            releaseLock.Set();
            holder?.Join(TimeSpan.FromSeconds(10));
            Capture(manager.Dispose);
            Capture(() => Directory.Delete(directory, recursive: true));
        }
        if (failure is not null)
            ExceptionDispatchInfo.Capture(failure).Throw();

        void Capture(Action action)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                failure = failure is null ? ex : new AggregateException(failure, ex);
            }
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Exit_LeavesAttachedWorkbookOpenWithoutSaving(bool save)
    {
        var directory = Path.Join(Path.GetTempPath(), $"CliServiceTrayExitTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        var owner = new SessionManager();
        var manager = new SessionManager();
        IExcelBatch? ownerBatch = null;
        string? name = null;
        Exception? failure = null;
        try
        {
            var ownerId = owner.CreateSessionForNewFile(Path.Join(directory, "owner.xlsx"));
            ownerBatch = Assert.IsAssignableFrom<IExcelBatch>(owner.GetSession(ownerId));
            ownerBatch.Execute((context, _) =>
            {
                Excel.Workbooks? books = null;
                Excel.Workbook? book = null;
                Excel.Sheets? sheets = null;
                Excel.Worksheet? sheet = null;
                Excel.Range? cell = null;
                try
                {
                    books = context.App.Workbooks;
                    book = books.Add();
                    name = book.Name;
                    sheets = book.Worksheets;
                    sheet = (Excel.Worksheet)sheets[1];
                    cell = sheet.Range["A1"];
                    cell.Value2 = "keep attached edits";
                }
                finally
                {
                    ComUtilities.Release(ref cell);
                    ComUtilities.Release(ref sheet);
                    ComUtilities.Release(ref sheets);
                    ComUtilities.Release(ref book);
                    ComUtilities.Release(ref books);
                }
            });
            var discovery = await manager.DiscoverOpenWorkbooksAsync();
            var selection = Assert.Single(discovery.Workbooks, book =>
                book.Name == name && book.ProcessId == ownerBatch.ExcelProcessId);
            manager.AttachWorkbook(selection.WorkbookId);
            manager.CreateSessionForNewFile(Path.Join(directory, "tray-owned.xlsx"));
            var prompts = new List<string>();
            var warnings = new List<string>();
            var shutdown = false;

            CliServiceTray.ExitService(manager, () =>
            {
                shutdown = true;
                return null;
            }, (message, _, _, _) =>
            {
                prompts.Add(message);
                return save ? DialogResult.Yes : DialogResult.No;
            }, (_, message, icon) =>
            {
                if (icon != ToolTipIcon.Info)
                    warnings.Add(message);
            });
            Assert.True(shutdown);
            Assert.Empty(warnings);
            manager.Dispose();

            ownerBatch.Execute((context, _) =>
            {
                Excel.Workbooks? books = null;
                Excel.Workbook? book = null;
                Excel.Sheets? sheets = null;
                Excel.Worksheet? sheet = null;
                Excel.Range? cell = null;
                try
                {
                    books = context.App.Workbooks;
                    Assert.Equal(2, books.Count);
                    book = books[name];
                    Assert.False(book.Saved);
                    Assert.Equal(string.Empty, book.Path);
                    sheets = book.Worksheets;
                    sheet = (Excel.Worksheet)sheets[1];
                    cell = sheet.Range["A1"];
                    Assert.Equal("keep attached edits", cell.Value2);
                }
                finally
                {
                    ComUtilities.Release(ref cell);
                    ComUtilities.Release(ref sheet);
                    ComUtilities.Release(ref sheets);
                    ComUtilities.Release(ref book);
                    ComUtilities.Release(ref books);
                }
            });
            Assert.Contains("1 active session(s)", Assert.Single(prompts), StringComparison.Ordinal);
        }
        catch (Exception ex)
        {
            failure = ex;
        }
        finally
        {
            Capture(() =>
            {
                if (ownerBatch is null || name is null)
                    return;
                ownerBatch.Execute((context, _) =>
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
                                break;
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
            });
            Capture(manager.Dispose);
            Capture(owner.Dispose);
            Capture(() => Directory.Delete(directory, recursive: true));
        }
        if (failure is not null)
            ExceptionDispatchInfo.Capture(failure).Throw();

        void Capture(Action action)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                failure = failure is null ? ex : new AggregateException(failure, ex);
            }
        }
    }
}
