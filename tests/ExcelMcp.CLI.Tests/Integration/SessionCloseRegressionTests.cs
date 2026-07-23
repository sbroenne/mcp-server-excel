using System.Collections.Concurrent;
using System.Reflection;
using System.Text.Json;
using Microsoft.Extensions.Logging.Abstractions;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "ServiceDaemon")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Medium")]
public sealed class SessionCloseRegressionTests : IClassFixture<TempDirectoryFixture>
{
    private readonly TempDirectoryFixture _fixture;

    public SessionCloseRegressionTests(TempDirectoryFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact(Timeout = 60000)]
    public async Task SessionClose_WhenDisposeFails_QuarantinesSessionAndRetryDoesNotReportAlreadyClosed()
    {
        using var service = new ExcelMcpService();
        var batch = new FakeBatch
        {
            WorkbookPath = CreateFakeWorkbookPath(),
            DisposeException = new InvalidOperationException("synthetic teardown failure")
        };
        const string sessionId = "dispose-failure-quarantine";
        RegisterSession(service, sessionId, batch, addKnownSessionId: true);

        var firstClose = await CloseSessionAsync(service, sessionId);

        Assert.False(firstClose.Success);
        Assert.NotNull(firstClose.ErrorMessage);
        Assert.Contains("Failed to dispose session", firstClose.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("synthetic teardown failure", firstClose.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        var listAfterFailedClose = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
        Assert.True(listAfterFailedClose.Success);
        Assert.NotNull(listAfterFailedClose.Result);
        Assert.Contains(sessionId, listAfterFailedClose.Result, StringComparison.Ordinal);

        var quarantinedUse = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.save",
            SessionId = sessionId
        });
        Assert.False(quarantinedUse.Success);
        Assert.NotNull(quarantinedUse.ErrorMessage);
        Assert.Contains("quarantined", quarantinedUse.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        var secondClose = await CloseSessionAsync(service, sessionId);

        Assert.False(secondClose.Success);
        var secondCloseText = (secondClose.ErrorMessage ?? string.Empty) + (secondClose.Result ?? string.Empty);
        Assert.DoesNotContain("already closed", secondCloseText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact(Timeout = 60000)]
    public async Task SessionClose_DuringInFlightOperation_ReturnsBusyAndKeepsSessionUsable()
    {
        using var service = new ExcelMcpService();
        var batch = new FakeBatch
        {
            WorkbookPath = CreateFakeWorkbookPath(),
            BlockSaveUntilReleased = true
        };
        const string sessionId = "in-flight-close-race";
        RegisterSession(service, sessionId, batch, addKnownSessionId: true);

        var inFlightSave = Task.Run(async () => await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.save",
            SessionId = sessionId
        }));
        Assert.True(await batch.WaitForSaveStartedAsync(TimeSpan.FromSeconds(10)));

        var closeWhileBusy = await CloseSessionAsync(service, sessionId);

        Assert.False(closeWhileBusy.Success);
        Assert.NotNull(closeWhileBusy.ErrorMessage);
        Assert.Contains("operation(s) still running", closeWhileBusy.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Wait for all operations to complete", closeWhileBusy.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        var listWhileBusy = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
        Assert.True(listWhileBusy.Success);
        Assert.NotNull(listWhileBusy.Result);
        using (var listJson = JsonDocument.Parse(listWhileBusy.Result))
        {
            var session = listJson.RootElement.GetProperty("sessions")
                .EnumerateArray()
                .Single(item => item.GetProperty("sessionId").GetString() == sessionId);
            Assert.Equal(1, session.GetProperty("activeOperations").GetInt32());
            Assert.False(session.GetProperty("canClose").GetBoolean());
        }

        batch.ReleaseSave();
        var saveResponse = await inFlightSave;
        Assert.True(saveResponse.Success);
        Assert.Equal(0, GetSessionManager(service).GetActiveOperationCount(sessionId));

        var secondSave = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.save",
            SessionId = sessionId
        });
        Assert.True(secondSave.Success);

        var finalClose = await CloseSessionAsync(service, sessionId);
        Assert.True(finalClose.Success);
    }

    private string CreateFakeWorkbookPath()
    {
        return Path.Combine(_fixture.TempDir, $"fake-batch-{Guid.NewGuid():N}.xlsx");
    }

    private static Task<ServiceResponse> CloseSessionAsync(ExcelMcpService service, string sessionId)
    {
        return service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new { save = false }, ServiceProtocol.JsonOptions)
        });
    }

    private static void RegisterSession(ExcelMcpService service, string sessionId, FakeBatch batch, bool addKnownSessionId)
    {
        var sessionManager = GetSessionManager(service);
        var activeSessions = GetPrivateField<ConcurrentDictionary<string, IExcelBatch>>(sessionManager, "_activeSessions");
        var activeFilePaths = GetPrivateField<ConcurrentDictionary<string, string>>(sessionManager, "_activeFilePaths");
        var sessionFilePaths = GetPrivateField<ConcurrentDictionary<string, string>>(sessionManager, "_sessionFilePaths");
        var activeOperationCounts = GetPrivateField<ConcurrentDictionary<string, int>>(sessionManager, "_activeOperationCounts");
        var showExcelFlags = GetPrivateField<ConcurrentDictionary<string, bool>>(sessionManager, "_showExcelFlags");
        var sessionOrigins = GetPrivateField<ConcurrentDictionary<string, SessionOrigin>>(sessionManager, "_sessionOrigins");
        var sessionCreatedAt = GetPrivateField<ConcurrentDictionary<string, DateTime>>(sessionManager, "_sessionCreatedAt");

        var normalizedPath = Path.GetFullPath(batch.WorkbookPath);
        activeSessions[sessionId] = batch;
        activeFilePaths[normalizedPath] = sessionId;
        sessionFilePaths[sessionId] = normalizedPath;
        activeOperationCounts[sessionId] = 0;
        showExcelFlags[sessionId] = false;
        sessionOrigins[sessionId] = SessionOrigin.CLI;
        sessionCreatedAt[sessionId] = DateTime.UtcNow;

        if (addKnownSessionId)
        {
            var knownSessionIds = GetPrivateField<ConcurrentDictionary<string, byte>>(service, "_knownSessionIds");
            knownSessionIds[sessionId] = 0;
        }
    }

    private static SessionManager GetSessionManager(ExcelMcpService service)
    {
        return GetPrivateField<SessionManager>(service, "_sessionManager");
    }

    private static T GetPrivateField<T>(object instance, string fieldName)
    {
        var field = instance.GetType().GetField(fieldName, BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(field);
        return (T)field!.GetValue(instance)!;
    }

    private sealed class FakeBatch : IExcelBatch
    {
        private readonly TaskCompletionSource _saveStarted = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource _releaseSave = new(TaskCreationOptions.RunContinuationsAsynchronously);

        public string WorkbookPath { get; init; } = string.Empty;
        public Microsoft.Extensions.Logging.ILogger Logger { get; } = NullLogger.Instance;
        public IReadOnlyDictionary<string, Excel.Workbook> Workbooks { get; } = new Dictionary<string, Excel.Workbook>();
        public bool HasTimedOutOperation => false;
        public int? ExcelProcessId => 1234;
        public TimeSpan OperationTimeout => TimeSpan.FromSeconds(5);
        public bool BlockSaveUntilReleased { get; init; }
        public Exception? DisposeException { get; init; }
        public int SaveCalls { get; private set; }
        public int DisposeCalls { get; private set; }

        public Excel.Workbook GetWorkbook(string filePath) => throw new NotSupportedException();

        public void Execute(Action<ExcelContext, CancellationToken> operation, CancellationToken cancellationToken = default)
        {
            throw new NotSupportedException();
        }

        public T Execute<T>(Func<ExcelContext, CancellationToken, T> operation, CancellationToken cancellationToken = default)
        {
            throw new NotSupportedException();
        }

        public void Save(CancellationToken cancellationToken = default)
        {
            SaveCalls++;
            if (!BlockSaveUntilReleased)
            {
                return;
            }

            _saveStarted.TrySetResult();
            _releaseSave.Task.Wait(cancellationToken);
        }

        public bool IsExcelProcessAlive() => true;

        public async Task<bool> WaitForSaveStartedAsync(TimeSpan timeout)
        {
            var completed = await Task.WhenAny(_saveStarted.Task, Task.Delay(timeout));
            return completed == _saveStarted.Task;
        }

        public void ReleaseSave()
        {
            _releaseSave.TrySetResult();
        }

        public void Dispose()
        {
            DisposeCalls++;
            if (DisposeException != null)
            {
                throw DisposeException;
            }
        }
    }
}
