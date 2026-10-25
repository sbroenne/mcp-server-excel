using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    [Fact]
    public async Task Shutdown_EditDuringSavedWorkbookClose_StopsAndLeavesChanges()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var id = await CreateSessionAsync(service, Path.Join(directory, "edit-during-shutdown.xlsx"));
            sessions[id] = 0;
            await CreateSheetAsync(service, id, "DeleteAfterRefusal");
            RequireSuccess(await service.ProcessAsync(new ServiceRequest
            {
                Command = "range.set-values",
                SessionId = id,
                Args = """{"sheetName":"DeleteAfterRefusal","rangeAddress":"A1","values":[["nonempty sheet"]],"overwritePolicy":"allow"}"""
            }));
            RequireSuccess(await service.ProcessAsync(new ServiceRequest { Command = "session.save", SessionId = id }));
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(id));
            var alertsEnabled = false;
            Excel.AppEvents_WorkbookBeforeCloseEventHandler? beforeClose = null;
            batch.Execute((context, _) =>
            {
                Assert.True(context.Book.Saved);
                Assert.False(context.App.DisplayAlerts);
                beforeClose = (Excel.Workbook workbook, ref bool cancel) =>
                {
                    cancel = true;
                    context.App.WorkbookBeforeClose -= beforeClose;
                    alertsEnabled = context.App.DisplayAlerts;
                    Excel.Sheets? sheets = null;
                    Excel.Worksheet? sheet = null;
                    Excel.Range? cell = null;
                    try
                    {
                        sheets = workbook.Worksheets;
                        sheet = (Excel.Worksheet)sheets["Sheet1"];
                        cell = sheet.Range["A1"];
                        cell.Value2 = "edit made during close";
                    }
                    finally
                    {
                        ComUtilities.Release(ref cell);
                        ComUtilities.Release(ref sheet);
                        ComUtilities.Release(ref sheets);
                    }
                };
                context.App.WorkbookBeforeClose += beforeClose;
            });

            var shutdown = await service.ProcessAsync(new ServiceRequest { Command = "service.shutdown" });
            Assert.True(shutdown.Success, shutdown.ErrorMessage);
            Assert.Contains("did not close", shutdown.PartialState, StringComparison.OrdinalIgnoreCase);
            Assert.True(alertsEnabled);
            Assert.Same(batch, service.SessionManager.GetSession(id));
            Assert.Equal("SessionUnavailable", (await service.ProcessAsync(new ServiceRequest
            {
                Command = "range.get-values",
                SessionId = id,
                Args = """{"sheetName":"Sheet1","rangeAddress":"A1"}"""
            })).ErrorCategory);
            batch.Execute((context, _) =>
            {
                Excel.Sheets? sheets = null;
                Excel.Worksheet? sheet = null;
                Excel.Range? cell = null;
                try
                {
                    Assert.False(context.Book.Saved);
                    sheets = context.Book.Worksheets;
                    Assert.Equal(2, sheets.Count);
                    sheet = (Excel.Worksheet)sheets["Sheet1"];
                    cell = sheet.Range["A1"];
                    Assert.Equal("edit made during close", ((object)cell.Value2));
                }
                finally
                {
                    ComUtilities.Release(ref cell);
                    ComUtilities.Release(ref sheet);
                    ComUtilities.Release(ref sheets);
                }
            });
        });
    }

    [Fact]
    public async Task Shutdown_CallbackCompletesAfterWaitExpires_StopsAndReportsThenHandsOver()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "late-shutdown.xlsx");
            var id = GetSessionId(await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.create",
                Args = JsonSerializer.Serialize(new { filePath = path, timeoutSeconds = 10 })
            }));
            sessions[id] = 0;
            await WriteMarkerAsync(service, id, "unsaved after expired shutdown wait");
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(id));
            var identity = Assert.IsType<ExcelProcessIdentity>(batch.ProcessIdentity);
            using var release = new ManualResetEventSlim();
            var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            batch.BeforeRefreshStateReadHookForTests = () =>
            {
                entered.TrySetResult();
                Assert.True(release.Wait(TimeSpan.FromSeconds(60)));
            };
            var shutdown = Task.Run(() => service.ProcessAsync(new ServiceRequest { Command = "service.shutdown" }));
            try
            {
                await entered.Task.WaitAsync(TimeSpan.FromSeconds(15));
                var expired = await shutdown.WaitAsync(TimeSpan.FromSeconds(20));
                Assert.True(expired.Success, expired.ErrorMessage);
                Assert.Contains("wait expired", expired.PartialState, StringComparison.OrdinalIgnoreCase);
                Assert.Contains($"Excel process {identity.ProcessId}", expired.PartialState, StringComparison.Ordinal);
                Assert.False(((IExcelBatchCloseState)batch).ShutdownCommitted);
                Assert.Same(batch, service.SessionManager.GetSession(id));
                Assert.False(SessionManager.IsExcelProcessPreserved(identity));

                batch.BeforeRefreshStateReadHookForTests = null;
                release.Set();
                Assert.True(SpinWait.SpinUntil(
                    () => ((IExcelBatchCloseState)batch).ShutdownCommitted
                        && !batch.OperationStatus.ExecutionPending,
                    TimeSpan.FromSeconds(15)));
                service.SessionManager.Dispose();
                sessions.TryRemove(id, out _);
                Assert.Equal(0, service.SessionManager.ActiveSessionCount);
                Assert.False(service.SessionManager.HasActiveConnections);
                Assert.True(SessionManager.IsExcelProcessPreserved(identity));
                Assert.True(OwnedProcessGuard.MatchesLiveProcess(identity));
            }
            finally
            {
                batch.BeforeRefreshStateReadHookForTests = null;
                release.Set();
                await shutdown.WaitAsync(TimeSpan.FromSeconds(20));
                batch.Dispose();
                using var cleanup = new ExcelMcpService();
                var attachedId = await AttachHandoverAsync(cleanup, path);
                Assert.Equal("unsaved after expired shutdown wait", await ReadMarkerAsync(cleanup, attachedId));
                await CloseSessionAsync(cleanup, attachedId, save: false);
            }
        });
    }

    [Fact]
    public async Task Shutdown_HandsOverUnsavedWorkbookWithoutSavingAndClosesSavedWorkbook()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "handover.xlsx");
            var id = await CreateSessionAsync(service, path);
            sessions[id] = 0;
            await WriteMarkerAsync(service, id, "saved baseline");
            RequireSuccess(await service.ProcessAsync(new ServiceRequest { Command = "session.save", SessionId = id }));
            await WriteMarkerAsync(service, id, "unsaved handover");
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(id));
            var savedId = await CreateSessionAsync(service, Path.Join(directory, "saved.xlsx"));
            sessions[savedId] = 0;
            var savedBatch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(savedId));
            var handedOver = false;
            try
            {
                RequireSuccess(await service.ProcessAsync(new ServiceRequest { Command = "service.shutdown" }));
                handedOver = true;
                sessions.TryRemove(id, out _);
                sessions.TryRemove(savedId, out _);
                service.Dispose();
                service.Dispose();
                Assert.True(batch.IsExcelProcessAlive());
                Assert.False(savedBatch.IsExcelProcessAlive());

                using var replacement = new ExcelMcpService();
                var attachedId = await AttachHandoverAsync(replacement, path);
                Assert.Equal("unsaved handover", await ReadMarkerAsync(replacement, attachedId));
                var attached = replacement.SessionManager.GetSession(attachedId)!;
                attached.Execute((context, _) =>
                {
                    Assert.False(context.Book.Saved);
                    Assert.True(context.App.Visible);
                    Assert.True(context.App.DisplayAlerts);
                });
                await CloseSessionAsync(replacement, attachedId, save: false);
                var reopened = await OpenSessionAsync(replacement, path);
                Assert.Equal("saved baseline", await ReadMarkerAsync(replacement, reopened));
                await CloseSessionAsync(replacement, reopened, save: false);
            }
            finally
            {
                if (handedOver && batch.IsExcelProcessAlive())
                {
                    using var cleanup = new ExcelMcpService();
                    var cleanupId = await AttachHandoverAsync(cleanup, path);
                    await CloseSessionAsync(cleanup, cleanupId, save: false);
                }
            }
        });
    }

    private static async Task<string> AttachHandoverAsync(
        ExcelMcpService service, string path, int? processId = null)
    {
        var discovery = await service.ProcessAsync(new ServiceRequest { Command = "session.discover" });
        RequireSuccess(discovery);
        using var json = JsonDocument.Parse(discovery.Result!);
        var workbook = Assert.Single(json.RootElement.GetProperty("workbooks").EnumerateArray(),
            item => item.TryGetProperty("location", out var location)
                && string.Equals(location.GetString(), path, StringComparison.OrdinalIgnoreCase)
                && (!processId.HasValue
                    || !item.TryGetProperty("processId", out var discoveredProcessId)
                    || discoveredProcessId.GetInt32() == processId.Value));
        return GetSessionId(await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.attach",
            Args = JsonSerializer.Serialize(new { workbookId = workbook.GetProperty("workbookId").GetString() })
        }));
    }
}
