using System.Collections.Concurrent;
using System.Reflection;
using System.Text.Json;
using Microsoft.Extensions.Logging.Abstractions;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Service;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "ExcelMcpService")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ExcelMcpServiceOperationTrackingTests
{
    [Fact]
    public async Task CompletedRead_ResponseCompletesAfterTrackingEnds_AndImmediateListAndCloseSucceed()
    {
        using var service = new ExcelMcpService();
        var batch = new FakeReadBatch();
        const string sessionId = "completed-read";
        RegisterSession(service, sessionId, batch);
        var manager = service.SessionManager;

        var readTask = service.ProcessAsync(CreateReadRequest(sessionId));
        var countWhenResponseCompletes = readTask.ContinueWith(
            _ => manager.GetActiveOperationCount(sessionId),
            CancellationToken.None,
            TaskContinuationOptions.ExecuteSynchronously,
            TaskScheduler.Default);

        var response = await readTask;

        Assert.True(response.Success);
        Assert.Equal(0, await countWhenResponseCompletes);
        AssertSessionState(await ListSessionsAsync(service), sessionId, 0, canClose: true);
        Assert.True((await CloseSessionAsync(service, sessionId)).Success);
    }

    [Fact]
    public async Task FailedRead_ResponseCompletesAfterTrackingEnds_AndImmediateCloseSucceeds()
    {
        using var service = new ExcelMcpService();
        var batch = new FakeReadBatch
        {
            ExecuteException = new InvalidOperationException("synthetic read failure")
        };
        const string sessionId = "failed-read";
        RegisterSession(service, sessionId, batch);

        var response = await service.ProcessAsync(CreateReadRequest(sessionId));

        Assert.False(response.Success);
        Assert.Contains("synthetic read failure", response.ErrorMessage, StringComparison.Ordinal);
        AssertSessionState(await ListSessionsAsync(service), sessionId, 0, canClose: true);
        Assert.True((await CloseSessionAsync(service, sessionId)).Success);
    }

    [Fact]
    public async Task CancelledRead_RemovesSessionAndOperationTrackingBeforeResponseCompletes()
    {
        using var service = new ExcelMcpService();
        var batch = new FakeReadBatch
        {
            ExecuteException = new OperationCanceledException("synthetic read cancellation")
        };
        const string sessionId = "cancelled-read";
        RegisterSession(service, sessionId, batch);

        var response = await service.ProcessAsync(CreateReadRequest(sessionId));

        Assert.False(response.Success);
        Assert.Equal("Cancelled", response.ErrorCategory);
        Assert.Equal(0, service.SessionManager.GetActiveOperationCount(sessionId));
        AssertSessionAbsent(await ListSessionsAsync(service), sessionId);
        Assert.True((await CloseSessionAsync(service, sessionId)).Success);
    }

    [Fact]
    public async Task NestedReads_CountEachActiveCallAndReturnToZero()
    {
        using var service = new ExcelMcpService();
        var batch = new FakeReadBatch();
        const string sessionId = "nested-reads";
        RegisterSession(service, sessionId, batch);
        var manager = service.SessionManager;
        ServiceResponse? nestedResponse = null;
        var nestedActiveCount = 0;

        batch.OnExecute = callNumber =>
        {
            if (callNumber == 1)
            {
                nestedResponse = service.ProcessAsync(CreateReadRequest(sessionId))
                    .GetAwaiter()
                    .GetResult();
            }
            else
            {
                nestedActiveCount = manager.GetActiveOperationCount(sessionId);
            }
        };

        var response = await service.ProcessAsync(CreateReadRequest(sessionId));

        Assert.True(response.Success);
        Assert.NotNull(nestedResponse);
        Assert.True(nestedResponse.Success);
        Assert.Equal(2, nestedActiveCount);
        Assert.Equal(0, manager.GetActiveOperationCount(sessionId));
        Assert.True((await CloseSessionAsync(service, sessionId)).Success);
    }

    [Fact]
    public async Task ConcurrentReads_KeepCloseBlockedOnlyUntilBothResponsesComplete()
    {
        using var service = new ExcelMcpService();
        var batch = new FakeReadBatch(expectedBlockedCalls: 2);
        const string sessionId = "concurrent-reads";
        RegisterSession(service, sessionId, batch);

        var firstRead = Task.Run(() => service.ProcessAsync(CreateReadRequest(sessionId)));
        var secondRead = Task.Run(() => service.ProcessAsync(CreateReadRequest(sessionId)));
        var bothReadsStarted = batch.WaitForBlockedCalls(TimeSpan.FromSeconds(10));
        ServiceResponse? listWhileBusy = null;
        ServiceResponse? busyClose = null;
        try
        {
            if (bothReadsStarted)
            {
                listWhileBusy = await ListSessionsAsync(service);
                busyClose = await CloseSessionAsync(service, sessionId);
            }
        }
        finally
        {
            batch.ReleaseBlockedCalls();
        }

        var responses = await Task.WhenAll(firstRead, secondRead);

        Assert.True(bothReadsStarted);
        AssertSessionState(listWhileBusy!, sessionId, 2, canClose: false);
        Assert.NotNull(busyClose);
        Assert.False(busyClose.Success);
        Assert.Contains("2 operation(s) still running", busyClose.ErrorMessage, StringComparison.Ordinal);
        Assert.All(responses, response => Assert.True(response.Success));
        AssertSessionState(await ListSessionsAsync(service), sessionId, 0, canClose: true);
        Assert.True((await CloseSessionAsync(service, sessionId)).Success);
    }

    private static ServiceRequest CreateReadRequest(string sessionId) => new()
    {
        Command = "range.get-values",
        SessionId = sessionId,
        Args = """{"sheetName":"Sheet1","rangeAddress":"A1"}"""
    };

    private static Task<ServiceResponse> ListSessionsAsync(ExcelMcpService service) =>
        service.ProcessAsync(new ServiceRequest { Command = "session.list" });

    private static Task<ServiceResponse> CloseSessionAsync(ExcelMcpService service, string sessionId) =>
        service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = sessionId,
            Args = """{"save":false}"""
        });

    private static void AssertSessionState(
        ServiceResponse response,
        string sessionId,
        int activeOperations,
        bool canClose)
    {
        Assert.True(response.Success);
        Assert.NotNull(response.Result);
        using var document = JsonDocument.Parse(response.Result);
        var session = document.RootElement.GetProperty("sessions")
            .EnumerateArray()
            .Single(item => item.GetProperty("sessionId").GetString() == sessionId);
        Assert.Equal(activeOperations, session.GetProperty("activeOperations").GetInt32());
        Assert.Equal(canClose, session.GetProperty("canClose").GetBoolean());
    }

    private static void AssertSessionAbsent(ServiceResponse response, string sessionId)
    {
        Assert.True(response.Success);
        Assert.NotNull(response.Result);
        using var document = JsonDocument.Parse(response.Result);
        Assert.DoesNotContain(
            document.RootElement.GetProperty("sessions").EnumerateArray(),
            item => item.GetProperty("sessionId").GetString() == sessionId);
    }

    private static void RegisterSession(
        ExcelMcpService service,
        string sessionId,
        FakeReadBatch batch)
    {
        var manager = service.SessionManager;
        var normalizedPath = Path.GetFullPath(batch.WorkbookPath);
        GetPrivateField<ConcurrentDictionary<string, IExcelBatch>>(manager, "_activeSessions")[sessionId] = batch;
        GetPrivateField<ConcurrentDictionary<string, string>>(manager, "_activeFilePaths")[normalizedPath] = sessionId;
        GetPrivateField<ConcurrentDictionary<string, string>>(manager, "_sessionFilePaths")[sessionId] = normalizedPath;
        GetPrivateField<ConcurrentDictionary<string, int>>(manager, "_activeOperationCounts")[sessionId] = 0;
        GetPrivateField<ConcurrentDictionary<string, bool>>(manager, "_showExcelFlags")[sessionId] = false;
        GetPrivateField<ConcurrentDictionary<string, SessionOrigin>>(manager, "_sessionOrigins")[sessionId] = SessionOrigin.CLI;
        GetPrivateField<ConcurrentDictionary<string, DateTime>>(manager, "_sessionCreatedAt")[sessionId] = DateTime.UtcNow;
        GetPrivateField<ConcurrentDictionary<string, byte>>(service, "_knownSessionIds")[sessionId] = 0;
    }

    private static T GetPrivateField<T>(object instance, string fieldName)
    {
        var field = instance.GetType().GetField(fieldName, BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(field);
        return (T)field.GetValue(instance)!;
    }

    private sealed class FakeReadBatch : IExcelBatch
    {
        private readonly CountdownEvent? _blockedCallsStarted;
        private readonly ManualResetEventSlim _releaseBlockedCalls = new(initialState: false);
        private int _executeCalls;

        public FakeReadBatch(int expectedBlockedCalls = 0)
        {
            if (expectedBlockedCalls > 0)
            {
                _blockedCallsStarted = new CountdownEvent(expectedBlockedCalls);
            }
        }

        public string WorkbookPath { get; } =
            Path.Combine(Path.GetTempPath(), $"operation-tracking-{Guid.NewGuid():N}.xlsx");
        public Microsoft.Extensions.Logging.ILogger Logger { get; } = NullLogger.Instance;
        public IReadOnlyDictionary<string, Excel.Workbook> Workbooks { get; } =
            new Dictionary<string, Excel.Workbook>();
        public bool HasTimedOutOperation => false;
        public int? ExcelProcessId => 1234;
        public TimeSpan OperationTimeout => TimeSpan.FromSeconds(5);
        public bool IsExcelVisible => false;
        public Exception? ExecuteException { get; init; }
        public Action<int>? OnExecute { get; set; }

        public Excel.Workbook GetWorkbook(string filePath) => throw new NotSupportedException();

        public void UpdateWorkbookPath(string workbookPath) => throw new NotSupportedException();

        public void Execute(
            Action<ExcelContext, CancellationToken> operation,
            CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();

        public T Execute<T>(
            Func<ExcelContext, CancellationToken, T> operation,
            CancellationToken cancellationToken = default)
        {
            var callNumber = Interlocked.Increment(ref _executeCalls);
            OnExecute?.Invoke(callNumber);
            if (_blockedCallsStarted != null)
            {
                _blockedCallsStarted.Signal();
                _releaseBlockedCalls.Wait(cancellationToken);
            }

            if (ExecuteException != null)
            {
                throw ExecuteException;
            }

            Assert.Equal(typeof(RangeValueResult), typeof(T));
            return (T)(object)new RangeValueResult
            {
                Success = true,
                FilePath = WorkbookPath,
                SheetName = "Sheet1",
                RangeAddress = "$A$1",
                RowCount = 1,
                ColumnCount = 1,
                Values = [["value"]]
            };
        }

        public void Save(CancellationToken cancellationToken = default)
        {
        }

        public bool IsExcelProcessAlive() => true;

        public bool WaitForBlockedCalls(TimeSpan timeout) =>
            _blockedCallsStarted?.Wait(timeout) ?? true;

        public void ReleaseBlockedCalls() => _releaseBlockedCalls.Set();

        public void Dispose()
        {
            _blockedCallsStarted?.Dispose();
            _releaseBlockedCalls.Dispose();
        }
    }
}
