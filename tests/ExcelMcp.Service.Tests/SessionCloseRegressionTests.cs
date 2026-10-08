using System.Collections.Concurrent;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using Microsoft.Extensions.Logging.Abstractions;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
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
    [Theory]
    [InlineData("rpc")]
    [InlineData("direct")]
    [InlineData("dispose")]
    public async Task Shutdown_FinalSaveBusyHResult_RetainsSessionAndAllowsRetry(string caller)
    {
        using var service = new ExcelMcpService();
        var comFailure = Assert.IsType<COMException>(Marshal.GetExceptionForHR(unchecked((int)0x800AC472)));
        var batch = new FakeBatch
        {
            WorkbookPath = CreateFakeWorkbookPath(),
            SaveException = ExcelShutdownService.CreateSaveFailureException(comFailure, "fixture.xlsx")
        };
        const string session = "final-save-busy";
        RegisterSession(service, session, batch, addKnownSessionId: true);
        var host = GetPrivateField<Sbroenne.ExcelMcp.Service.Rpc.DaemonHost>(service, "_daemonHost");
        var shutdown = GetPrivateField<CancellationTokenSource>(host, "_shutdownCts");
        var idleCount = GetPrivateField<Func<int>>(host, "_sessionCount");
        batch.BeforeSave = () => Assert.Equal(1, idleCount());
        try
        {
            if (caller == "rpc")
            {
                var response = await service.ProcessAsync(new ServiceRequest { Command = "service.shutdown" });
                Assert.False(response.Success);
                Assert.Equal("Busy", response.ErrorCategory);
                Assert.Equal("0x800AC472", response.HResult);
            }
            else
            {
                var failure = Assert.Throws<AggregateException>(
                    caller == "direct" ? service.RequestShutdown : service.Dispose);
                var busy = Assert.IsType<ExcelBusyException>(Assert.Single(failure.InnerExceptions).InnerException);
                Assert.Same(comFailure, busy.InnerException);
            }
            Assert.Equal(0, batch.DisposeCalls);
            Assert.Same(batch, service.SessionManager.GetSession(session));
            Assert.True(service.SessionManager.TryGetFilePath(session, out var retainedPath));
            Assert.Equal(batch.WorkbookPath, retainedPath);
            Assert.False(shutdown.IsCancellationRequested);
            Assert.True((await service.ProcessAsync(new ServiceRequest { Command = "service.ping" })).Success);
        }
        finally { batch.SaveException = null; }

        var retry = await service.ProcessAsync(new ServiceRequest { Command = "service.shutdown" });
        Assert.True(retry.Success, retry.ErrorMessage);
        Assert.Equal(1, batch.DisposeCalls);
        Assert.Equal(0, service.SessionManager.ActiveSessionCount);
    }

    [Fact]
    public void DaemonIdleCount_RemovesDeadSessionsWithoutProbingLiveExcelReadiness()
    {
        using var service = new ExcelMcpService();
        var dead = new FakeBatch { WorkbookPath = CreateFakeWorkbookPath() };
        var live = new FakeBatch
        {
            WorkbookPath = CreateFakeWorkbookPath(),
            RefreshState = WorkbookRefreshState.Unknown
        };
        RegisterSession(service, "dead-for-idle", dead, addKnownSessionId: true);
        RegisterSession(service, "live-for-idle", live, addKnownSessionId: true);
        var host = GetPrivateField<Sbroenne.ExcelMcp.Service.Rpc.DaemonHost>(service, "_daemonHost");
        var count = GetPrivateField<Func<int>>(host, "_sessionCount");
        try
        {
            Assert.Equal(2, count());
            dead.IsAlive = false;
            Assert.Equal(1, count());
            Assert.Equal(1, dead.DisposeCalls);
            Assert.False(service.SessionManager.TryGetFilePath("dead-for-idle", out _));
            Assert.Same(live, service.SessionManager.GetSession("live-for-idle"));
            Assert.Equal(0, live.RefreshProbeCalls);
            live.IsAlive = false;
            Assert.Equal(0, count());
            Assert.Equal(1, live.DisposeCalls);
            Assert.Equal(0, service.SessionManager.ActiveSessionCount);
        }
        finally { live.RefreshState = WorkbookRefreshState.Ready; }
    }

    [Theory]
    [InlineData("rpc")]
    [InlineData("direct")]
    [InlineData("dispose")]
    public async Task Shutdown_ConcurrentCaller_CannotBypassBusyRefusal(string caller)
    {
        using var service = new ExcelMcpService();
        using var saveEntered = new ManualResetEventSlim();
        using var releaseSave = new ManualResetEventSlim();
        using var secondStarted = new ManualResetEventSlim();
        var saveAttempts = 0;
        var batch = new FakeBatch
        {
            WorkbookPath = CreateFakeWorkbookPath(),
            RefreshState = WorkbookRefreshState.Refreshing,
            BeforeSave = () =>
            {
                if (Interlocked.Increment(ref saveAttempts) != 1) return;
                saveEntered.Set();
                Assert.True(releaseSave.Wait(TimeSpan.FromSeconds(10)));
            }
        };
        const string session = "concurrent-shutdown-busy";
        RegisterSession(service, session, batch, addKnownSessionId: true);
        var host = GetPrivateField<Sbroenne.ExcelMcp.Service.Rpc.DaemonHost>(service, "_daemonHost");
        var shutdown = GetPrivateField<CancellationTokenSource>(host, "_shutdownCts");
        var first = Task.Run(() => service.ProcessAsync(new ServiceRequest { Command = "service.shutdown" }));
        Task<(ServiceResponse? Response, Exception? Error)>? second = null;
        try
        {
            Assert.True(saveEntered.Wait(TimeSpan.FromSeconds(10)));
            second = Task.Run(async () =>
            {
                secondStarted.Set();
                if (caller == "rpc")
                    return (await service.ProcessAsync(new ServiceRequest { Command = "service.shutdown" }), (Exception?)null);
                return ((ServiceResponse?)null,
                    Record.Exception(caller == "direct" ? service.RequestShutdown : service.Dispose));
            });
            Assert.True(secondStarted.Wait(TimeSpan.FromSeconds(10)));
            await Task.WhenAny(second, Task.Delay(250));
            releaseSave.Set();

            var refused = await first.WaitAsync(TimeSpan.FromSeconds(10));
            Assert.False(refused.Success);
            Assert.Equal("Busy", refused.ErrorCategory);
            var concurrent = await second.WaitAsync(TimeSpan.FromSeconds(10));
            if (caller == "rpc")
            {
                Assert.NotNull(concurrent.Response);
                Assert.False(concurrent.Response.Success);
                Assert.Equal("Busy", concurrent.Response.ErrorCategory);
            }
            else
            {
                Assert.IsType<AggregateException>(concurrent.Error);
            }
            Assert.False(shutdown.IsCancellationRequested);
            Assert.Equal(0, batch.DisposeCalls);
            Assert.Same(batch, service.SessionManager.GetSession(session));
            Assert.True((await service.ProcessAsync(new ServiceRequest { Command = "service.ping" })).Success);
        }
        finally
        {
            releaseSave.Set();
            await first.WaitAsync(TimeSpan.FromSeconds(10));
            if (second != null) await second.WaitAsync(TimeSpan.FromSeconds(10));
            batch.RefreshState = WorkbookRefreshState.Ready;
        }
        var retry = await service.ProcessAsync(new ServiceRequest { Command = "service.shutdown" });
        Assert.True(retry.Success, retry.ErrorMessage);
        Assert.Equal(1, batch.DisposeCalls);
        Assert.Equal(0, service.SessionManager.ActiveSessionCount);
    }

    [Theory]
    [InlineData("get-account-settings")]
    [InlineData("clear-account-hint")]
    [InlineData("set-account-settings")]
    public async Task ConnectionAccountSettings_RefreshStartsAfterPreflight_ReturnsBusy(string action)
    {
        using var service = new ExcelMcpService();
        var batch = new FakeBatch
        {
            WorkbookPath = CreateFakeWorkbookPath(),
            ExecuteException = Record.Exception(() =>
                typeof(ConnectionCommands).GetMethod(
                    "ValidateAccountSettingsConnectionReadiness",
                    BindingFlags.Static | BindingFlags.NonPublic)!.Invoke(null, [true]))
        };
        Assert.NotNull(batch.ExecuteException);
        const string session = "account-refresh-race";
        RegisterSession(service, session, batch, addKnownSessionId: true);
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "connection." + action,
            SessionId = session,
            Args = action == "set-account-settings"
                ? """{"connectionName":"Selected","accountHint":"Synthetic hint"}"""
                : """{"connectionName":"Selected"}"""
        });
        Assert.False(response.Success);
        Assert.Equal("Busy", response.ErrorCategory);
        Assert.True(batch.ExecuteCalls > 0);
        Assert.Equal(0, batch.DisposeCalls);
        Assert.Same(batch, service.SessionManager.GetSession(session));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public async Task Shutdown_NonReadySession_RefusesAndRetainsServiceForRetry(int state)
    {
        using var service = new ExcelMcpService();
        var batch = new FakeBatch
        {
            WorkbookPath = CreateFakeWorkbookPath(),
            RefreshState = (WorkbookRefreshState)state
        };
        const string session = "shutdown-busy";
        RegisterSession(service, session, batch, addKnownSessionId: true);
        try
        {
            var response = await service.ProcessAsync(new ServiceRequest { Command = "service.shutdown" });
            Assert.False(response.Success);
            Assert.Equal("Busy", response.ErrorCategory);
            Assert.Equal(0, batch.DisposeCalls);
            Assert.Same(batch, service.SessionManager.GetSession(session));
            Assert.True((await service.ProcessAsync(new ServiceRequest { Command = "service.ping" })).Success);
        }
        finally
        {
            batch.RefreshState = WorkbookRefreshState.Ready;
        }
        var retry = await service.ProcessAsync(new ServiceRequest { Command = "service.shutdown" });
        Assert.True(retry.Success, retry.ErrorMessage);
        Assert.Equal(1, batch.DisposeCalls);
        Assert.Equal(0, service.SessionManager.ActiveSessionCount);
    }

    [Theory]
    [InlineData("get-refresh-status", 2)]
    [InlineData("get-refresh-status", 3)]
    [InlineData("get-refresh-status", 4)]
    [InlineData("cancel-refresh", 2)]
    [InlineData("cancel-refresh", 3)]
    [InlineData("cancel-refresh", 4)]
    public async Task ConnectionRefreshControls_RejectUnavailableExcelBeforeQueuingCom(string action, int state)
    {
        using var service = new ExcelMcpService();
        var batch = new FakeBatch
        {
            WorkbookPath = CreateFakeWorkbookPath(),
            RefreshState = (WorkbookRefreshState)state
        };
        const string session = "refresh-controls-busy";
        RegisterSession(service, session, batch, addKnownSessionId: true);
        try
        {
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "connection." + action,
                SessionId = session,
                Args = """{"connectionName":"Selected"}"""
            });
            Assert.False(response.Success);
            Assert.Equal("Busy", response.ErrorCategory);
            Assert.Equal(0, batch.ExecuteCalls);
            Assert.Same(batch, service.SessionManager.GetSession(session));
        }
        finally
        {
            batch.RefreshState = WorkbookRefreshState.Ready;
        }
    }

    [Theory]
    [InlineData("get-account-settings", 1)]
    [InlineData("get-account-settings", 2)]
    [InlineData("get-account-settings", 3)]
    [InlineData("get-account-settings", 4)]
    [InlineData("clear-account-hint", 1)]
    [InlineData("clear-account-hint", 2)]
    [InlineData("clear-account-hint", 3)]
    [InlineData("clear-account-hint", 4)]
    [InlineData("set-account-settings", 1)]
    [InlineData("set-account-settings", 2)]
    [InlineData("set-account-settings", 3)]
    [InlineData("set-account-settings", 4)]
    public async Task ConnectionAccountSettings_RejectsNonReadyExcelBeforeQueuingCom(string action, int state)
    {
        using var service = new ExcelMcpService();
        var batch = new FakeBatch
        {
            WorkbookPath = CreateFakeWorkbookPath(),
            RefreshState = (WorkbookRefreshState)state
        };
        const string session = "account-settings-busy";
        RegisterSession(service, session, batch, addKnownSessionId: true);
        try
        {
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "connection." + action,
                SessionId = session,
                Args = """{"connectionName":"Selected"}"""
            });
            Assert.False(response.Success);
            Assert.Equal("Busy", response.ErrorCategory);
            Assert.Contains("account settings", response.ErrorMessage, StringComparison.Ordinal);
            Assert.Equal(0, batch.ExecuteCalls);
            Assert.Same(batch, GetSessionManager(service).GetSession(session));
        }
        finally
        {
            batch.RefreshState = WorkbookRefreshState.Ready;
        }
    }

    [Theory]
    [InlineData(false, 0)]
    [InlineData(true, 0)]
    [InlineData(false, 1)]
    [InlineData(true, 1)]
    public async Task DialogOpen_ReportsActionableStateAndRetainsSession(bool save, int activeOperations)
    {
        using var service = new ExcelMcpService();
        var batch = new FakeBatch
        {
            WorkbookPath = CreateFakeWorkbookPath(),
            RefreshState = WorkbookRefreshState.DialogOpen
        };
        const string sessionId = "dialog-open";
        RegisterSession(service, sessionId, batch, addKnownSessionId: true);
        var operationCounts = GetPrivateField<ConcurrentDictionary<string, int>>(
            GetSessionManager(service), "_activeOperationCounts");
        operationCounts[sessionId] = activeOperations;

        try
        {
            var list = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
            Assert.True(list.Success, list.ErrorMessage);
            using var json = JsonDocument.Parse(list.Result!);
            var session = Assert.Single(json.RootElement.GetProperty("sessions").EnumerateArray());
            Assert.Equal("dialogOpen", session.GetProperty("excelState").GetString());
            Assert.Contains("Check the Excel window", session.GetProperty("blockingReason").GetString(), StringComparison.Ordinal);
            Assert.Equal(activeOperations, session.GetProperty("activeOperations").GetInt32());
            Assert.False(session.GetProperty("canClose").GetBoolean());

            var closed = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.close",
                SessionId = sessionId,
                Args = JsonSerializer.Serialize(new { save }, ServiceProtocol.JsonOptions)
            });
            Assert.False(closed.Success);
            Assert.Equal("Busy", closed.ErrorCategory);
            Assert.Contains("dialog", closed.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("sign-in required", closed.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, batch.DisposeCalls);
            Assert.Same(batch, GetSessionManager(service).GetSession(sessionId));
        }
        finally
        {
            batch.RefreshState = WorkbookRefreshState.Ready;
            operationCounts[sessionId] = 0;
        }

        var readyList = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
        using var readyJson = JsonDocument.Parse(readyList.Result!);
        var readySession = Assert.Single(readyJson.RootElement.GetProperty("sessions").EnumerateArray());
        Assert.Equal("ready", readySession.GetProperty("excelState").GetString());
        Assert.True(readySession.GetProperty("canClose").GetBoolean());
        Assert.False(readySession.TryGetProperty("blockingReason", out _));
        var retried = await CloseSessionAsync(service, sessionId);
        Assert.True(retried.Success, retried.ErrorMessage);
        Assert.Equal(1, batch.DisposeCalls);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public async Task NativeRefreshWithoutTrackedOperation_PreventsCloseAndRetainsSession(int state)
    {
        using var service = new ExcelMcpService();
        var batch = new FakeBatch
        {
            WorkbookPath = CreateFakeWorkbookPath(),
            RefreshState = (WorkbookRefreshState)state
        };
        const string sessionId = "native-refresh";
        RegisterSession(service, sessionId, batch, addKnownSessionId: true);

        var list = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
        Assert.True(list.Success, list.ErrorMessage);
        using (var json = JsonDocument.Parse(list.Result!))
        {
            var session = Assert.Single(json.RootElement.GetProperty("sessions").EnumerateArray());
            Assert.Equal(0, session.GetProperty("activeOperations").GetInt32());
            Assert.False(session.GetProperty("canClose").GetBoolean());
            Assert.Equal(batch.RefreshState.ToString().ToLowerInvariant(), session.GetProperty("excelState").GetString());
            Assert.Contains("refresh", session.GetProperty("blockingReason").GetString(), StringComparison.OrdinalIgnoreCase);
        }
        var closed = await CloseSessionAsync(service, sessionId);
        Assert.False(closed.Success);
        Assert.Equal("Busy", closed.ErrorCategory);
        Assert.Contains("refresh", closed.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Same(batch, GetSessionManager(service).GetSession(sessionId));
        Assert.Equal(0, batch.DisposeCalls);

        batch.RefreshState = WorkbookRefreshState.Ready;
        var retried = await CloseSessionAsync(service, sessionId);
        Assert.True(retried.Success, retried.ErrorMessage);
        Assert.Equal(1, batch.DisposeCalls);
    }

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

    private sealed class FakeBatch : IExcelBatch, IExcelBatchRefreshState
    {
        public WorkbookRefreshState RefreshState { get; set; } = WorkbookRefreshState.Ready;
        public int RefreshProbeCalls { get; private set; }
        public WorkbookRefreshState GetRefreshState()
        {
            RefreshProbeCalls++;
            return RefreshState;
        }
        public string WorkbookPath { get; init; } = string.Empty;
        public Microsoft.Extensions.Logging.ILogger Logger { get; } = NullLogger.Instance;
        public IReadOnlyDictionary<string, Excel.Workbook> Workbooks { get; } = new Dictionary<string, Excel.Workbook>();
        public bool HasTimedOutOperation => false;
        public int? ExcelProcessId => 1234;
        public TimeSpan OperationTimeout => TimeSpan.FromSeconds(5);
        public bool IsExcelVisible => false;
        public Exception? DisposeException { get; init; }
        public Exception? ExecuteException { get; init; }
        public Action? BeforeSave { get; set; }
        public Exception? SaveException { get; set; }
        public bool IsAlive { get; set; } = true;
        public int DisposeCalls { get; private set; }
        public int ExecuteCalls { get; private set; }
        public bool BlockOperations { get; init; }
        public ManualResetEventSlim OperationEntered { get; } = new();
        public ManualResetEventSlim ReleaseOperation { get; } = new();

        public Excel.Workbook GetWorkbook(string filePath) => throw new NotSupportedException();

        public void UpdateWorkbookPath(string workbookPath) => throw new NotSupportedException();

        public void Execute(Action<ExcelContext, CancellationToken> operation, CancellationToken cancellationToken = default)
        {
            ExecuteCalls++;
            WaitForRelease(cancellationToken);
            if (ExecuteException != null) throw ExecuteException;
            throw new NotSupportedException();
        }

        public T Execute<T>(Func<ExcelContext, CancellationToken, T> operation, CancellationToken cancellationToken = default)
        {
            ExecuteCalls++;
            WaitForRelease(cancellationToken);
            if (ExecuteException != null) throw ExecuteException;
            throw new NotSupportedException();
        }

        public void Save(CancellationToken cancellationToken = default)
        {
            BeforeSave?.Invoke();
            ExcelBusyException.ThrowIfNotReady(RefreshState, "save");
            if (SaveException != null) throw SaveException;
        }

        public bool IsExcelProcessAlive() => IsAlive;

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
