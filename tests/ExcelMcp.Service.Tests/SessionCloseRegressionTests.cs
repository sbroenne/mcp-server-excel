using System.Collections.Concurrent;
using System.Reflection;
using System.Text.Json;
using Microsoft.Extensions.Logging.Abstractions;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "ServiceDaemon")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Medium")]
public sealed class SessionCloseRegressionTests
{
    [Fact]
    public async Task SessionClose_MissingSessionReturnsStructuredError()
    {
        using var service = new ExcelMcpService();

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = "missing-session",
            Args = """{"save":false}"""
        });

        Assert.False(response.Success);
        Assert.Equal("SessionNotFound", response.ErrorCategory);
        Assert.Equal("missing-session", response.SessionId);
        Assert.Contains("not found", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact(Timeout = 60000)]
    public async Task SessionClose_WhenDisposeFails_QuarantinesSessionAndRetryDoesNotReportAlreadyClosed()
    {
        using var service = new ExcelMcpService();
        var batch = new FakeBatch
        {
            WorkbookPath = CreateFakeWorkbookPath(),
            DisposeException = new InvalidOperationException(
                "Excel process 4321 did not exit and remains tracked for pipe cleanup")
        };
        const string sessionId = "dispose-failure-quarantine";
        RegisterSession(service, sessionId, batch, addKnownSessionId: true);

        var firstClose = await CloseSessionAsync(service, sessionId);

        Assert.False(firstClose.Success);
        Assert.NotNull(firstClose.ErrorMessage);
        Assert.Contains("Failed to dispose session", firstClose.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("remains tracked", firstClose.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        var listAfterFailedClose = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
        Assert.True(listAfterFailedClose.Success);
        Assert.NotNull(listAfterFailedClose.Result);
        Assert.Contains(sessionId, listAfterFailedClose.Result, StringComparison.Ordinal);

        var quarantinedUse = await service.ProcessAsync(new ServiceRequest
        {
            Command = "sheet.list",
            SessionId = sessionId
        });
        Assert.False(quarantinedUse.Success);
        Assert.NotNull(quarantinedUse.ErrorMessage);
        Assert.Contains("quarantined", quarantinedUse.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        var secondClose = await CloseSessionAsync(service, sessionId);

        Assert.False(secondClose.Success);
        var secondCloseText = (secondClose.ErrorMessage ?? string.Empty) + (secondClose.Result ?? string.Empty);
        Assert.DoesNotContain("already closed", secondCloseText, StringComparison.OrdinalIgnoreCase);

        var shutdownFailure = Assert.Throws<AggregateException>(service.Dispose);
        var sessionFailure = Assert.Single(shutdownFailure.InnerExceptions);
        Assert.Contains(sessionId, sessionFailure.Message, StringComparison.Ordinal);
        Assert.Same(batch.DisposeException, sessionFailure.InnerException);
        Assert.Throws<ObjectDisposedException>(service.RequestShutdown);
    }

    [Fact(Timeout = 60000)]
    public async Task SessionClose_DuringInFlightOperation_ReturnsBusyAndKeepsSessionUsable()
    {
        using var service = new ExcelMcpService();
        var batch = new FakeBatch
        {
            WorkbookPath = CreateFakeWorkbookPath(),
            BlockOperations = true
        };
        const string sessionId = "in-flight-close-race";
        RegisterSession(service, sessionId, batch, addKnownSessionId: true);

        var operationTask = Task.Run(() => service.ProcessAsync(new ServiceRequest
        {
            Command = "sheet.list",
            SessionId = sessionId
        }));
        Assert.True(
            batch.OperationEntered.Wait(TimeSpan.FromSeconds(5)),
            "The Service operation did not enter the registered batch.");

        ServiceResponse? operationResponse = null;
        try
        {
            var closeWhileBusy = await CloseSessionAsync(service, sessionId);

            Assert.False(closeWhileBusy.Success);
            Assert.NotNull(closeWhileBusy.ErrorMessage);
            Assert.Contains("operation(s) still running", closeWhileBusy.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Wait for all operations to complete", closeWhileBusy.ErrorMessage, StringComparison.OrdinalIgnoreCase);

            var listWhileBusy = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
            Assert.True(listWhileBusy.Success);
            Assert.NotNull(listWhileBusy.Result);
            using var listJson = JsonDocument.Parse(listWhileBusy.Result);
            var session = listJson.RootElement.GetProperty("sessions")
                .EnumerateArray()
                .Single(item => item.GetProperty("sessionId").GetString() == sessionId);
            Assert.Equal(1, session.GetProperty("activeOperations").GetInt32());
            Assert.False(session.GetProperty("canClose").GetBoolean());
        }
        finally
        {
            batch.ReleaseOperation.Set();
            operationResponse = await operationTask;
        }

        Assert.False(operationResponse.Success);
        Assert.Equal(nameof(NotSupportedException), operationResponse.ExceptionType);
        Assert.Contains("not supported", operationResponse.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        var listAfterOperation = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
        Assert.True(listAfterOperation.Success, listAfterOperation.ErrorMessage);
        Assert.NotNull(listAfterOperation.Result);
        using (var listJson = JsonDocument.Parse(listAfterOperation.Result))
        {
            var session = listJson.RootElement.GetProperty("sessions")
                .EnumerateArray()
                .Single(item => item.GetProperty("sessionId").GetString() == sessionId);
            Assert.Equal(0, session.GetProperty("activeOperations").GetInt32());
            Assert.True(session.GetProperty("canClose").GetBoolean());
        }

        var finalClose = await CloseSessionAsync(service, sessionId);
        Assert.True(finalClose.Success);
    }

    private static string CreateFakeWorkbookPath()
    {
        return Path.Combine(Path.GetTempPath(), $"fake-batch-{Guid.NewGuid():N}.xlsx");
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
        public string WorkbookPath { get; init; } = string.Empty;
        public Microsoft.Extensions.Logging.ILogger Logger { get; } = NullLogger.Instance;
        public IReadOnlyDictionary<string, Excel.Workbook> Workbooks { get; } = new Dictionary<string, Excel.Workbook>();
        public bool HasTimedOutOperation => false;
        public int? ExcelProcessId => 1234;
        public TimeSpan OperationTimeout => TimeSpan.FromSeconds(5);
        public bool IsExcelVisible => false;
        public Exception? DisposeException { get; init; }
        public int DisposeCalls { get; private set; }
        public bool BlockOperations { get; init; }
        public ManualResetEventSlim OperationEntered { get; } = new();
        public ManualResetEventSlim ReleaseOperation { get; } = new();

        public Excel.Workbook GetWorkbook(string filePath) => throw new NotSupportedException();

        public void UpdateWorkbookPath(string workbookPath) => throw new NotSupportedException();

        public void Execute(Action<ExcelContext, CancellationToken> operation, CancellationToken cancellationToken = default)
        {
            WaitForRelease(cancellationToken);
            throw new NotSupportedException();
        }

        public T Execute<T>(Func<ExcelContext, CancellationToken, T> operation, CancellationToken cancellationToken = default)
        {
            WaitForRelease(cancellationToken);
            throw new NotSupportedException();
        }

        public void Save(CancellationToken cancellationToken = default)
        {
        }

        public bool IsExcelProcessAlive() => true;

        public void Dispose()
        {
            DisposeCalls++;
            ReleaseOperation.Set();
            if (DisposeException != null)
            {
                throw DisposeException;
            }
        }

        private void WaitForRelease(CancellationToken cancellationToken)
        {
            if (!BlockOperations)
            {
                return;
            }

            OperationEntered.Set();
            ReleaseOperation.Wait(cancellationToken);
        }
    }
}
