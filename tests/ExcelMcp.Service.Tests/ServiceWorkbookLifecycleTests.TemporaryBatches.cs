using System.Diagnostics;
using System.Runtime.ExceptionServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Service.Workers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    [Theory]
    [InlineData(false, false, false, false, false)]
    [InlineData(true, false, false, false, false)]
    [InlineData(false, true, false, false, false)]
    [InlineData(true, true, false, false, false)]
    [InlineData(false, true, true, false, false)]
    [InlineData(false, true, false, true, false)]
    [InlineData(false, true, false, true, true)]
    public async Task TemporaryBatch_CompletedCrossFileCall_ClosesOnlyOnSuccess(
        bool move, bool failSave, bool readOnlySource, bool renameTarget, bool usePathAliases)
    {
        var directory = Path.GetFullPath(Path.Join(AppContext.BaseDirectory,
            "..", "..", "..", "..", "..", "artifacts", $"tb-{Guid.NewGuid():N}"));
        Directory.CreateDirectory(directory);
        using var processes = new TestExcelProcessScope();
        var service = new ExcelMcpService();
        var source = Path.Join(directory, "source.xlsx");
        var failure = await Record.ExceptionAsync(async () =>
        {
            var target = Path.Join(directory, "target.xlsx");
            var session = await CreateSessionAsync(service, source);
            await WriteMarkerAsync(service, session, "Completed transfer");
            await CreateSheetAsync(service, session, "Keep");
            await CloseSessionAsync(service, session, save: true);
            session = await CreateSessionAsync(service, target);
            RequireSuccess(await service.ProcessAsync(new ServiceRequest
            {
                Command = "sheet.rename",
                SessionId = session,
                Args = JsonSerializer.Serialize(new { oldName = "Sheet1", newName = "Destination" })
            }));
            await CloseSessionAsync(service, session, save: true);
            if (readOnlySource)
            {
                ExcelSession.TemporaryBatchCreatedForTests = batch =>
                {
                    batch.WaitForStartup();
                    batch.Execute((_, _) => batch.GetWorkbook(source).ChangeFileAccess(Excel.XlFileAccess.xlReadOnly));
                };
            }
            if (failSave)
                ExcelShutdownService.BeforeSaveHookForTests = () =>
                    throw new InvalidOperationException("Original temporary save failure.");
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = move ? "sheet.move-to-file" : "sheet.copy-to-file",
                Args = JsonSerializer.Serialize(new
                {
                    sourceFile = usePathAliases ? Path.Join(directory, ".", "source.xlsx") : source,
                    targetFile = usePathAliases ? Path.Join(directory, ".", "target.xlsx") : target,
                    sourceSheet = "Sheet1"
                })
            });
            ExcelShutdownService.BeforeSaveHookForTests = null;
            if (failSave)
            {
                Assert.False(response.Success);
                Assert.Equal("InvalidOperationException", response.ExceptionType);
                Assert.Contains("Original temporary save failure.", response.ErrorMessage, StringComparison.Ordinal);
                var retained = Assert.Single(service.SessionManager.ActiveSessionIds);
                Assert.Equal(retained, response.SessionId);
                using var partial = JsonDocument.Parse(response.PartialState!);
                Assert.True(partial.RootElement.GetProperty("retained").GetBoolean());
                Assert.False(partial.RootElement.GetProperty("operation").GetProperty("executionPending").GetBoolean());
                Assert.False(partial.RootElement.GetProperty("operation").GetProperty("waitExpired").GetBoolean());
                if (renameTarget)
                {
                    var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(retained));
                    var renamed = Path.Join(directory, "target-renamed.xlsx");
                    batch.Execute((_, _) => batch.GetWorkbook(target).SaveAs(renamed));
                    var discovery = await service.SessionManager.DiscoverOpenWorkbooksAsync(batch.ProcessIdentity);
                    Assert.True(discovery.Complete, string.Join("; ", discovery.Errors.Select(error => error.Message)));
                    Assert.Contains(discovery.Workbooks, workbook => workbook.Location == renamed);
                    Assert.True(service.SessionManager.TryGetFilePath(retained, out var selected));
                    Assert.Equal(source, selected);
                    Assert.Equal(source, batch.WorkbookPath);
                    batch.Execute((context, _) =>
                    {
                        Assert.Equal(source, context.Book.FullName);
                        Assert.Equal(renamed, batch.GetWorkbook(renamed).FullName);
                    });
                    Assert.Equal(new[] { source, renamed }, service.SessionManager.GetOwnedFilePaths(retained));
                    var listed = await service.ProcessAsync(new ServiceRequest
                    {
                        Command = "session.list"
                    });
                    RequireSuccess(listed);
                    using var sessions = JsonDocument.Parse(listed.Result!);
                    Assert.Equal(new[] { source, renamed }, Assert.Single(
                        sessions.RootElement.GetProperty("sessions").EnumerateArray())
                        .GetProperty("filePaths").EnumerateArray().Select(path => path.GetString()));
                    var sheetResponse = await service.ProcessAsync(new ServiceRequest
                    {
                        Command = "sheet.list",
                        SessionId = retained,
                        Args = JsonSerializer.Serialize(new { filePath = renamed })
                    });
                    RequireSuccess(sheetResponse);
                    using var worksheets = JsonDocument.Parse(sheetResponse.Result!);
                    Assert.Equal(2, worksheets.RootElement.GetProperty("worksheets").GetArrayLength());
                    Assert.Throws<KeyNotFoundException>(() => batch.Execute((_, _) => batch.GetWorkbook(target)));
                    target = renamed;
                }
                if (readOnlySource)
                {
                    var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(retained));
                    batch.Execute((_, _) =>
                    {
                        Assert.True(batch.GetWorkbook(source).ReadOnly);
                        Assert.True(batch.GetWorkbook(source).Saved);
                        Assert.False(batch.GetWorkbook(target).Saved);
                    });
                }
                RequireSuccess(await service.ProcessAsync(new ServiceRequest { Command = "session.save", SessionId = retained }));
                await CloseSessionAsync(service, retained, save: false);
            }
            else
                RequireSuccess(response);
            Assert.Empty(service.SessionManager.ActiveSessionIds);
            var reopened = await OpenSessionAsync(service, target);
            Assert.Equal("Completed transfer", await ReadMarkerAsync(service, reopened));
            await CloseSessionAsync(service, reopened, save: false);
            reopened = await OpenSessionAsync(service, source);
            var list = await service.ProcessAsync(new ServiceRequest { Command = "sheet.list", SessionId = reopened });
            RequireSuccess(list);
            using var sheets = JsonDocument.Parse(list.Result!);
            Assert.Equal(move ? 1 : 2, sheets.RootElement.GetProperty("worksheets").GetArrayLength());
            await CloseSessionAsync(service, reopened, save: false);
        });
        ExcelShutdownService.BeforeSaveHookForTests = null;
        ExcelSession.TemporaryBatchCreatedForTests = null;
        var cleanup = Record.Exception(() =>
        {
            foreach (var session in service.SessionManager.ActiveSessionIds)
                service.SessionManager.CloseSession(session, save: false);
            service.Dispose();
            if (processes.FindLeftovers() is { } leftover)
                throw leftover;
            Directory.Delete(directory, recursive: true);
        });
        if (failure is not null && cleanup is not null)
            throw new AggregateException(failure, cleanup);
        if (failure is not null || cleanup is not null)
            ExceptionDispatchInfo.Capture((failure ?? cleanup)!).Throw();
    }

    [Theory]
    [InlineData(false, false, false, false)]
    [InlineData(true, false, false, false)]
    [InlineData(false, true, false, false)]
    [InlineData(true, true, false, false)]
    [InlineData(false, false, true, false)]
    [InlineData(true, false, true, false)]
    [InlineData(false, false, false, true)]
    public async Task TemporaryBatch_CrossFileWaitExpires_RetainsWorkerAndExactWorkbooks(
        bool move, bool startup, bool failSave, bool loseController)
    {
        var directory = Path.GetFullPath(Path.Join(AppContext.BaseDirectory,
            "..", "..", "..", "..", "..", "artifacts", $"tb-{Guid.NewGuid():N}"));
        Directory.CreateDirectory(directory);
        using var processes = new TestExcelProcessScope();
        var service = new ExcelMcpService();
        using var controller = loseController ? Process.Start(new ProcessStartInfo("pwsh")
        {
            ArgumentList = { "-NoProfile", "-Command", "[Console]::ReadLine() | Out-Null" },
            RedirectStandardInput = true,
            UseShellExecute = false,
            CreateNoWindow = true
        })! : null;
        var registration = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}",
            WorkerProcess.Current, controller is null ? WorkerProcess.Current : WorkerProcess.From(controller),
            Guid.NewGuid().ToString("N"),
            WorkerRegistry.Build, null);
        using var host = new ExcelWorkerHost(registration, service);
        var running = Task.CompletedTask;
        var worker = new ExcelWorkerClient(registration, WorkerProcess.Current, registration.Generation);
        using var release = new ManualResetEventSlim();
        var saving = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        ExcelBatch? temporary = null;
        Task<ServiceResponse>? operation = null;
        var failure = await Record.ExceptionAsync(async () =>
        {
            var source = Path.Join(directory, "source.xlsx");
            var target = Path.Join(directory, "target.xlsx");
            var session = await CreateSessionAsync(service, source);
            await WriteMarkerAsync(service, session, "Source marker");
            await CreateSheetAsync(service, session, "Keep");
            RequireSuccess(await service.ProcessAsync(new ServiceRequest
            {
                Command = "sheet.rename",
                SessionId = session,
                Args = JsonSerializer.Serialize(new { oldName = "Sheet1", newName = "Transferred" })
            }));
            await CloseSessionAsync(service, session, save: true);
            session = await CreateSessionAsync(service, target);
            await WriteMarkerAsync(service, session, "Target marker");
            await CloseSessionAsync(service, session, save: true);
            var sourceBytes = await File.ReadAllBytesAsync(source);
            var targetBytes = await File.ReadAllBytesAsync(target);
            running = host.RunAsync();
            if (!loseController)
                await worker.InitializeAsync();
            ExcelSession.TemporaryBatchTimeoutForTests = TimeSpan.FromSeconds(2);
            ExcelSession.TemporaryBatchStartupTimeoutForTests = startup ? TimeSpan.FromSeconds(2) : null;
            ExcelSession.TemporaryBatchCreatedForTests = batch => temporary = batch;
            void Block()
            {
                saving.TrySetResult();
                Assert.True(release.Wait(TimeSpan.FromSeconds(30)));
            }
            if (startup)
            {
                ExcelBatch.BeforeWorkbookOpenHook = (path, _) =>
                {
                    if (path == target)
                        Block();
                };
            }
            else
                ExcelShutdownService.BeforeSaveHookForTests = () =>
                {
                    Block();
                    if (failSave)
                        throw new InvalidOperationException("Cross-file save failed after worksheet transfer.");
                };
            var command = move ? "sheet.move-to-file" : "sheet.copy-to-file";
            var request = new ServiceRequest
            {
                Command = command,
                Args = JsonSerializer.Serialize(new
                {
                    sourceFile = source,
                    targetFile = target,
                    sourceSheet = "Transferred"
                })
            };
            operation = loseController
                ? Task.Run(() => service.ProcessAsync(request))
                : worker.ExecuteAsync(request, CancellationToken.None);
            await Task.WhenAny(saving.Task, operation).WaitAsync(TimeSpan.FromSeconds(15));
            if (operation.IsCompleted && !startup)
                Assert.Fail($"Cross-file callback did not reach save: {(await operation).ErrorMessage}");
            await saving.Task.WaitAsync(TimeSpan.FromSeconds(10));
            var response = await operation.WaitAsync(TimeSpan.FromSeconds(10));
            Assert.False(response.Success);
            if (controller is not null)
            {
                controller.StandardInput.Close();
                await controller.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
                await Task.Delay(TimeSpan.FromSeconds(5));
                Assert.False(running.IsCompleted, "Controller loss must not classify the retained temporary batch as empty.");
                Assert.Single(service.SessionManager.ActiveSessionIds);
                await worker.InitializeAsync();
            }
            var released = await worker.ReleaseAsync(CancellationToken.None);
            Assert.False(released.Success, "Worker release must refuse while a temporary native callback is pending.");
            Assert.Equal("Busy", released.ErrorCategory);
            Assert.Equal("Timeout", response.ErrorCategory);
            Assert.Equal(startup ? "ExcelStartupPendingException" : "TimeoutException", response.ExceptionType);
            Assert.Contains("timed out", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Cannot dispose", response.ErrorMessage, StringComparison.Ordinal);
            Assert.Equal(command, response.Command);
            var retained = Assert.Single(service.SessionManager.ActiveSessionIds);
            Assert.Equal(retained, response.SessionId);
            var snapshot = await worker.ReadSnapshotAsync();
            Assert.Equal(new[] { source, target }, snapshot.OwnershipKeys[retained]);
            using (var partial = JsonDocument.Parse(response.PartialState!))
            {
                Assert.True(partial.RootElement.GetProperty("retained").GetBoolean());
                Assert.Equal(retained, partial.RootElement.GetProperty("sessionId").GetString());
                Assert.Equal(new[] { source, target },
                    partial.RootElement.GetProperty("filePaths").EnumerateArray().Select(path => path.GetString()));
                Assert.True(partial.RootElement.GetProperty("operation").GetProperty("executionPending").GetBoolean());
                if (!startup)
                    Assert.Equal(command, partial.RootElement.GetProperty("operation").GetProperty("command").GetString());
            }
            Assert.Equal(sourceBytes, ReadSharedWorkbookBytes(source));
            Assert.Equal(targetBytes, ReadSharedWorkbookBytes(target));
            var status = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
            RequireSuccess(status);
            using (var listed = JsonDocument.Parse(status.Result!))
            {
                var entry = Assert.Single(listed.RootElement.GetProperty("sessions").EnumerateArray());
                Assert.Equal(new[] { source, target },
                    entry.GetProperty("filePaths").EnumerateArray().Select(path => path.GetString()));
            }
            Assert.Same(temporary, service.SessionManager.GetSession(retained));
            Assert.True(service.SessionManager.GetOperationStatus(retained)!.ExecutionPending);
            Assert.True(service.SessionManager.GetOperationStatus(retained)!.WaitExpired);
            Assert.NotNull(temporary);
            using (var excel = Process.GetProcessById(temporary.ExcelProcessId!.Value))
                Assert.NotEqual(IntPtr.Zero, excel.MainWindowHandle);
            var record = Assert.Single(new WorkerRegistry().Read(), record => record.PipeName == registration.PipeName);
            Assert.Equal(temporary.ProcessIdentity, record.Excel);
            Assert.False(running.IsCompleted);
            Assert.False((await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.close",
                SessionId = retained
            })).Success);
            Assert.Same(temporary, service.SessionManager.GetSession(retained));
            ExcelShutdownService.BeforeSaveHookForTests = null;
            ExcelBatch.BeforeWorkbookOpenHook = null;
            release.Set();
            await WaitForTemporaryReadyAsync(temporary);
            if (failSave)
            {
                var failedOperation = service.SessionManager.GetOperationStatus(retained)!;
                Assert.Equal("failed", failedOperation.State);
                Assert.False(failedOperation.ExecutionPending);
                Assert.Contains("Cross-file save failed after worksheet transfer.", failedOperation.ErrorMessage,
                    StringComparison.Ordinal);
            }
            var state = temporary.Execute((_, _) =>
            {
                var sourceBook = temporary.GetWorkbook(source);
                var targetBook = temporary.GetWorkbook(target);
                Excel.Worksheet? transferred = null;
                Excel.Sheets? targetSheets = null;
                Excel.Sheets? sourceSheets = null;
                Excel.Range? cell = null;
                try
                {
                    targetSheets = targetBook.Worksheets;
                    sourceSheets = sourceBook.Worksheets;
                    transferred = (Excel.Worksheet)targetSheets[startup ? "Sheet1" : "Transferred"];
                    cell = transferred.Range["A1"];
                    return (SourceSheets: sourceSheets.Count, TargetSheets: targetSheets.Count,
                        Marker: (string)cell.Value2, SourceSaved: sourceBook.Saved, TargetSaved: targetBook.Saved);
                }
                finally
                {
                    ComUtilities.Release(ref cell);
                    ComUtilities.Release(ref transferred);
                    ComUtilities.Release(ref sourceSheets);
                    ComUtilities.Release(ref targetSheets);
                }
            });
            Assert.Equal(move && !startup ? 1 : 2, state.SourceSheets);
            Assert.Equal(startup ? 1 : 2, state.TargetSheets);
            Assert.Equal(startup ? "Target marker" : "Source marker", state.Marker);
            Assert.Equal(!move || !failSave, state.SourceSaved);
            Assert.Equal(!failSave, state.TargetSaved);
            if (failSave)
            {
                Assert.Equal(sourceBytes, ReadSharedWorkbookBytes(source));
                Assert.Equal(targetBytes, ReadSharedWorkbookBytes(target));
                if (!move)
                {
                    var discovery = await service.SessionManager.DiscoverOpenWorkbooksAsync(temporary.ProcessIdentity);
                    Assert.True(discovery.Complete, string.Join("; ", discovery.Errors.Select(error => error.Message)));
                    var secondary = Assert.Single(discovery.Workbooks, workbook => workbook.Location == target);
                    Assert.True(service.SessionManager.TryGetFilePath(retained, out var primary));
                    Assert.Equal(source, primary);
                    var attach = Assert.Throws<InvalidOperationException>(() =>
                        service.SessionManager.AttachWorkbook(secondary.WorkbookId));
                    Assert.Contains("already managed", attach.Message, StringComparison.OrdinalIgnoreCase);
                    Assert.Equal(new[] { source, target }, service.SessionManager.GetOwnedFilePaths(retained));
                }
                RequireSuccess(await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.save",
                    SessionId = retained
                }));
                var saved = temporary.Execute((_, _) =>
                    (Source: temporary.GetWorkbook(source).Saved, Target: temporary.GetWorkbook(target).Saved));
                Assert.True(saved.Source);
                Assert.True(saved.Target);
            }
            Assert.Same(temporary, service.SessionManager.GetSession(retained));
            await CloseSessionAsync(service, retained, save: false);
            Assert.Empty(service.SessionManager.ActiveSessionIds);
            RequireSuccess(await worker.ReleaseAsync(CancellationToken.None));
            await running.WaitAsync(TimeSpan.FromSeconds(15));
            using var verification = new ExcelMcpService();
            var reopened = await OpenSessionAsync(verification, target);
            Assert.Equal(startup ? "Target marker" : "Source marker",
                await ReadMarkerAsync(verification, reopened, startup ? "Sheet1" : "Transferred"));
            await CloseSessionAsync(verification, reopened, save: false);
        });
        ExcelSession.TemporaryBatchCreatedForTests = null;
        ExcelSession.TemporaryBatchTimeoutForTests = null;
        ExcelSession.TemporaryBatchStartupTimeoutForTests = null;
        ExcelShutdownService.BeforeSaveHookForTests = null;
        ExcelBatch.BeforeWorkbookOpenHook = null;
        release.Set();
        var cleanup = await Record.ExceptionAsync(async () =>
        {
            if (controller is { HasExited: false })
            {
                controller.StandardInput.Close();
                await controller.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
            }
            if (operation is not null)
                await operation.WaitAsync(TimeSpan.FromSeconds(15));
            if (temporary is not null && temporary.IsExcelProcessAlive())
            {
                await WaitForTemporaryReadyAsync(temporary);
                if (!service.SessionManager.ActiveSessionIds.Any(id =>
                    ReferenceEquals(service.SessionManager.GetSession(id), temporary)))
                    temporary.Close();
            }
            foreach (var session in service.SessionManager.ActiveSessionIds)
                service.SessionManager.CloseSession(session, save: false);
            if (!running.IsCompleted)
                RequireSuccess(await worker.ReleaseAsync(CancellationToken.None));
            await running.WaitAsync(TimeSpan.FromSeconds(15));
            service.Dispose();
            if (processes.FindLeftovers() is { } leftover)
                throw leftover;
            Directory.Delete(directory, recursive: true);
        });
        if (failure is not null && cleanup is not null)
            throw new AggregateException(failure, cleanup);
        if (failure is not null || cleanup is not null)
            ExceptionDispatchInfo.Capture((failure ?? cleanup)!).Throw();
    }

    private static async Task WaitForTemporaryReadyAsync(ExcelBatch batch)
    {
        var timer = Stopwatch.StartNew();
        while (batch.GetRefreshState() != WorkbookRefreshState.Ready)
        {
            Assert.True(timer.Elapsed < TimeSpan.FromSeconds(15), "Temporary Excel batch did not become ready.");
            await Task.Delay(50);
        }
    }

    private static byte[] ReadSharedWorkbookBytes(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        var bytes = new byte[stream.Length];
        stream.ReadExactly(bytes);
        return bytes;
    }
}
