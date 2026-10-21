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

            CliServiceTray.ExitService(manager, () => shutdown = true, (message, _, _, _) =>
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
