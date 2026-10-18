using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

/// <summary>
/// Regression tests for the ExcelBatch message pump.
///
/// These tests validate the fix for a critical bug where the STA thread message pump
/// degenerated into 100% CPU spin when idle. The bug had two independent mechanisms:
///
/// 1. Silent catch(Exception){} bypassed Thread.Sleep(10), causing 0ms-backoff spin
///    when any exception occurred (e.g., ObjectDisposedException on CancellationTokenSource).
///
/// 2. Thread.Sleep(10) on STA thread with registered OLE message filter returned immediately
///    when pending COM messages existed (Excel events during calculation), turning the poll
///    loop into a tight spin.
///
/// The fix replaced polling with WaitToReadAsync() which blocks efficiently and wakes
/// instantly when work arrives.
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test message pump CPU usage when idle (regression for 100% CPU spin)
/// - ✅ Test work processing after idle periods; report response times for diagnostics
/// - ✅ Test shutdown drains remaining work items (race condition fix)
/// - ✅ Test Dispose during Execute gives clean error (race condition fix)
///
/// IMPORTANT: These tests spawn and terminate Excel processes (side effects).
/// They run OnDemand only to avoid interference with normal test runs.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "ExcelBatch")]
[Trait("RunType", "OnDemand")]
[Collection("Sequential")] // Disable parallelization to avoid COM interference
[Trait("RequiresExcel", "true")]
public class ExcelBatchMessagePumpTests : IAsyncLifetime
{
    private readonly ITestOutputHelper _output;
    private static string? _staticTestFile;
    private string? _testFileCopy;

    public ExcelBatchMessagePumpTests(ITestOutputHelper output)
    {
        _output = output;
    }

    public Task InitializeAsync()
    {
        if (_staticTestFile == null)
        {
            var testFolder = Path.Join(AppContext.BaseDirectory, "Integration", "Session", "TestFiles");
            _staticTestFile = Path.Join(testFolder, "batch-test-static.xlsx");

            if (!File.Exists(_staticTestFile))
            {
                throw new FileNotFoundException($"Static test file not found at {_staticTestFile}. " +
                    "Please create the batch-test-static.xlsx file in the TestFiles folder.");
            }
        }

        _testFileCopy = Path.Join(Path.GetTempPath(), $"batch-pump-test-{Guid.NewGuid():N}.xlsx");
        File.Copy(_staticTestFile, _testFileCopy, overwrite: true);

        return Task.Delay(500);
    }

    public Task DisposeAsync()
    {
        if (_testFileCopy != null && File.Exists(_testFileCopy))
        {
            File.Delete(_testFileCopy);
        }
        return Task.CompletedTask;
    }

    /// <summary>
    /// REGRESSION TEST: Message pump must NOT spin at 100% CPU when idle.
    ///
    /// This is the primary regression test for the bug. It opens a batch, performs one
    /// operation to ensure everything is initialized, then measures CPU usage over a
    /// 3-second idle period. The original bug caused ~100% CPU on one core (2.97s/3s).
    /// The fix should show near-zero CPU.
    ///
    /// We measure CPU time of the EXCEL.EXE process (which should be ~0 since we're not
    /// doing anything) AND the overall thread behavior by measuring our own process's CPU.
    /// </summary>
    [Fact]
    public void MessagePump_WhenIdle_DoesNotSpinCpu()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);

        // Perform one operation to ensure everything is fully initialized
        Assert.Equal(_testFileCopy, batch.Execute((ctx, _) => ctx.Book.FullName));

        // Capture the Excel process ID for measurement
        int? excelPid = batch.ExcelProcessId;
        Assert.NotNull(excelPid);

        // Let everything settle
        Thread.Sleep(500);

        // Act — measure CPU over a 3-second idle window
        using var currentProcess = Process.GetCurrentProcess();
        var cpuBefore = currentProcess.TotalProcessorTime;
        var wallBefore = Stopwatch.GetTimestamp();

        // Also measure Excel's CPU
        TimeSpan excelCpuBefore;
        using (var excelProcess = Process.GetProcessById(excelPid.Value))
        {
            excelProcess.Refresh();
            excelCpuBefore = excelProcess.TotalProcessorTime;
        }

        // Idle period — the message pump should be sleeping, not spinning
        Thread.Sleep(3000);

        var cpuAfter = currentProcess.TotalProcessorTime;
        var wallAfter = Stopwatch.GetTimestamp();

        TimeSpan excelCpuAfter;
        using (var excelProcess = Process.GetProcessById(excelPid.Value))
        {
            excelProcess.Refresh();
            excelCpuAfter = excelProcess.TotalProcessorTime;
        }

        // Calculate
        var cpuUsed = (cpuAfter - cpuBefore).TotalMilliseconds;
        var wallElapsed = Stopwatch.GetElapsedTime(wallBefore, wallAfter).TotalMilliseconds;
        var cpuPercent = (cpuUsed / wallElapsed) * 100.0;

        var excelCpuUsed = (excelCpuAfter - excelCpuBefore).TotalMilliseconds;
        var excelCpuPercent = (excelCpuUsed / wallElapsed) * 100.0;

        _output.WriteLine($"Idle period: {wallElapsed:F0}ms wall time");
        _output.WriteLine($"MCP process CPU: {cpuUsed:F1}ms ({cpuPercent:F1}%)");
        _output.WriteLine($"Excel process CPU: {excelCpuUsed:F1}ms ({excelCpuPercent:F1}%)");

        // Assert — CPU should be well under 5% during idle.
        // The original bug showed ~100% (one full core). Even with test runner overhead,
        // a properly blocking message pump should use <1% CPU when idle.
        // Using 10% as threshold to avoid flaky failures from OS scheduling jitter,
        // background GC, or other test runner activity.
        Assert.True(cpuPercent < 10.0,
            $"REGRESSION: Message pump is spinning! CPU usage during 3s idle period: {cpuPercent:F1}% " +
            $"({cpuUsed:F1}ms CPU / {wallElapsed:F0}ms wall). " +
            "Expected <10%. The original bug showed ~100% CPU (one full core spinning).");

        _output.WriteLine($"✓ Message pump is idle — CPU usage: {cpuPercent:F1}% (threshold: <10%)");
    }

    /// <summary>
    /// Work submitted after idle periods must execute exactly once and return its result.
    /// Response times include COM guard overhead and machine scheduling, so they are
    /// diagnostic measurements rather than a message-pump latency requirement.
    /// </summary>
    [Fact]
    public void MessagePump_WhenWorkArrivesAfterIdle_ExecutesEachOperationOnce()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);

        // Warmup — ensure first-call JIT overhead is gone
        Assert.Equal(_testFileCopy, batch.Execute((ctx, _) => ctx.Book.FullName));

        // Let the pump return to idle state
        Thread.Sleep(200);

        // Act — measure round-trip time for a trivial operation
        // This measures: post to channel + pump wakes + execute + return result
        var latencies = new List<double>();
        var executedOperations = new List<int>();
        for (int i = 0; i < 20; i++)
        {
            // Wait between iterations so pump goes back to blocking wait
            Thread.Sleep(50);

            var operationId = i + 1;
            var sw = Stopwatch.StartNew();
            var result = batch.Execute((_, _) =>
            {
                executedOperations.Add(operationId);
                return operationId;
            });
            sw.Stop();
            Assert.Equal(operationId, result);
            latencies.Add(sw.Elapsed.TotalMilliseconds);
        }
        Assert.Equal(Enumerable.Range(1, 20), executedOperations);

        // Calculate statistics
        var sortedLatencies = latencies.OrderBy(x => x).ToList();
        var middle = sortedLatencies.Count / 2;
        var median = sortedLatencies.Count % 2 == 0
            ? (sortedLatencies[middle - 1] + sortedLatencies[middle]) / 2.0
            : sortedLatencies[middle];
        var p95 = sortedLatencies[(int)(sortedLatencies.Count * 0.95)];
        var max = sortedLatencies.Last();

        _output.WriteLine($"Execute round-trip diagnostics over {latencies.Count} iterations (includes COM guard overhead):");
        _output.WriteLine($"  Median: {median:F2}ms");
        _output.WriteLine($"  P95:    {p95:F2}ms");
        _output.WriteLine($"  Max:    {max:F2}ms");

    }

    /// <summary>
    /// RACE CONDITION TEST: Dispose with queued work must not drop or hang it.
    ///
    /// While one callback runs and another waits in the queue, Dispose() is refused
    /// and keeps Excel open. Both callers then receive their actual results, and a
    /// later Dispose() shuts Excel down.
    /// </summary>
    [Fact]
    public void Dispose_WithQueuedWork_RefusesAndCompletesQueuedWork()
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        Assert.Equal(_testFileCopy, batch.Execute((ctx, _) => ctx.Book.FullName));

        using var running = new ManualResetEventSlim(false);
        using var releaseRunning = new ManualResetEventSlim(false);
        using var queued = new ManualResetEventSlim(false);
        Exception? runningError = null, queuedError = null;
        int runningResult = -1, queuedResult = -1;

        var runningThread = new Thread(() =>
        {
            try
            {
                runningResult = batch.Execute((ctx, ct) =>
                {
                    running.Set();
                    Assert.True(releaseRunning.Wait(TimeSpan.FromSeconds(30)));
                    return 7;
                });
            }
            catch (Exception ex) { runningError = ex; }
        });
        var queuedThread = new Thread(() =>
        {
            try { queuedResult = batch.Execute((ctx, ct) => 42); }
            catch (Exception ex) { queuedError = ex; }
        });

        try
        {
            runningThread.Start();
            Assert.True(running.Wait(TimeSpan.FromSeconds(10)), "The first operation never started.");
            ExcelBatch.WorkItemQueuedHookForTests = queued.Set;
            queuedThread.Start();
            Assert.True(queued.Wait(TimeSpan.FromSeconds(10)), "The second operation was never queued.");

            var refused = Assert.Throws<ExcelBusyException>(batch.Dispose);
            Assert.Contains("still pending", refused.Message, StringComparison.Ordinal);
            Assert.True(batch.IsExcelProcessAlive(), "Refused dispose must keep Excel running.");
        }
        finally
        {
            ExcelBatch.WorkItemQueuedHookForTests = null;
            releaseRunning.Set();
            Assert.True(runningThread.Join(TimeSpan.FromSeconds(30)));
            Assert.True(queuedThread.Join(TimeSpan.FromSeconds(30)));
        }

        Assert.Null(runningError);
        Assert.Null(queuedError);
        Assert.Equal(7, runningResult);
        Assert.Equal(42, queuedResult);

        batch.Dispose();
        Assert.True(SpinWait.SpinUntil(() => !batch.IsExcelProcessAlive(), TimeSpan.FromSeconds(30)),
            "Excel must exit after queued work completes and dispose succeeds.");
    }

    /// <summary>
    /// RACE CONDITION TEST: Execute() after Dispose() gives clean ObjectDisposedException.
    ///
    /// When Dispose() completes the channel writer and a concurrent Execute() call passes
    /// the _disposed check but hits Writer.WriteAsync(), it should get a clean
    /// ObjectDisposedException — not an ugly ChannelClosedException.
    /// </summary>
    [Fact]
    public void Execute_AfterDispose_ThrowsObjectDisposedException()
    {
        // Arrange — create and immediately dispose
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);

        Assert.Equal(_testFileCopy, batch.Execute((ctx, _) => ctx.Book.FullName));

        batch.Dispose();

        // Wait for disposal to complete
        Thread.Sleep(500);

        // Act & Assert — Execute after Dispose should throw ObjectDisposedException
        var callbackRan = false;
        var ex = Assert.Throws<ObjectDisposedException>(() =>
        {
            batch.Execute((ctx, ct) => { callbackRan = true; return 0; });
        });
        Assert.False(callbackRan);

        _output.WriteLine($"✓ Got expected ObjectDisposedException: {ex.Message}");
    }

    /// <summary>
    /// RACE CONDITION TEST: Dispose during a running Execute is refused cleanly.
    ///
    /// Dispose() must not tear down Excel while a callback is running. It reports
    /// that work is pending, keeps Excel and the workbook open, lets the callback
    /// return its actual result, and succeeds once the work has finished.
    /// </summary>
    [Fact]
    public void Dispose_DuringActiveExecute_RefusesUntilWorkCompletes()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);

        // Initialize
        Assert.Equal(_testFileCopy, batch.Execute((ctx, _) => ctx.Book.FullName));

        // Start a long-running operation on another thread
        using var operationStarted = new ManualResetEventSlim(false);
        using var releaseOperation = new ManualResetEventSlim(false);
        Exception? executeError = null;
        int executeResult = -1;

        var executeThread = new Thread(() =>
        {
            try
            {
                executeResult = batch.Execute((ctx, ct) =>
                {
                    // Signal that we're inside the operation
                    operationStarted.Set();
                    // Simulate work — long enough for Dispose to race
                    Assert.True(releaseOperation.Wait(TimeSpan.FromSeconds(30)));
                    return 99;
                });
            }
            catch (Exception ex)
            {
                executeError = ex;
            }
        });

        // Act
        executeThread.Start();

        // Wait for the operation to actually start on the STA thread
        Assert.True(operationStarted.Wait(TimeSpan.FromSeconds(10)));
        _output.WriteLine("Operation started on STA thread, now calling Dispose...");

        // Call Dispose while the operation is mid-flight
        try
        {
            var refused = Assert.Throws<ExcelBusyException>(batch.Dispose);
            Assert.Contains("still pending", refused.Message, StringComparison.Ordinal);
            Assert.True(batch.IsExcelProcessAlive(), "Refused dispose must keep Excel running.");
            Assert.False(executeThread.Join(0), "Refused dispose must not end the running callback.");
        }
        finally
        {
            releaseOperation.Set();
            Assert.True(executeThread.Join(TimeSpan.FromSeconds(30)));
        }

        _output.WriteLine($"Execute result: {executeResult}");
        _output.WriteLine($"Execute error: {executeError?.GetType().Name}: {executeError?.Message}");

        Assert.Null(executeError);
        Assert.Equal(99, executeResult);
        Assert.Equal(_testFileCopy, batch.Execute((ctx, _) => ctx.Book.FullName));

        batch.Dispose();
        Assert.True(SpinWait.SpinUntil(() => !batch.IsExcelProcessAlive(), TimeSpan.FromSeconds(30)),
            "Excel must exit after the work completes and dispose succeeds.");
    }
}
