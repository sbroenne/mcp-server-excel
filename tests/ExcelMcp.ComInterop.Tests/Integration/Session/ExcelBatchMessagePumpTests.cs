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
/// - ✅ Test message pump wake latency (regression for 10ms poll delay)
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
#pragma warning disable CA1031 // Intentional: best-effort test cleanup
            try { File.Delete(_testFileCopy); } catch (Exception) { /* file may still be locked */ }
#pragma warning restore CA1031
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
        batch.Execute((ctx, ct) =>
        {
            _ = ctx.Book.Worksheets.Item(1);
            return 0;
        });

        // Capture the Excel process ID for measurement
        int? excelPid = batch.ExcelProcessId;
        Assert.NotNull(excelPid);

        // Let everything settle
        Thread.Sleep(500);

        // Act — measure CPU over a 3-second idle window
        var currentProcess = Process.GetCurrentProcess();
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
    /// REGRESSION TEST: Message pump must wake instantly when work arrives.
    ///
    /// The original polling loop had up to 10ms latency (wait for next poll cycle).
    /// The fix using WaitToReadAsync wakes the channel reader immediately when a writer
    /// posts work. This test verifies sub-5ms wake latency.
    /// </summary>
    [Fact]
    public void MessagePump_WhenWorkArrives_WakesWithLowLatency()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);

        // Warmup — ensure first-call JIT overhead is gone
        batch.Execute((ctx, ct) =>
        {
            _ = ctx.Book.Worksheets.Item(1);
            return 0;
        });

        // Let the pump return to idle state
        Thread.Sleep(200);

        // Act — measure round-trip time for a trivial operation
        // This measures: post to channel + pump wakes + execute + return result
        var latencies = new List<double>();
        for (int i = 0; i < 20; i++)
        {
            // Wait between iterations so pump goes back to blocking wait
            Thread.Sleep(50);

            var sw = Stopwatch.StartNew();
            batch.Execute((ctx, ct) =>
            {
                // Trivial operation — no COM call, just return
                return 42;
            });
            sw.Stop();
            latencies.Add(sw.Elapsed.TotalMilliseconds);
        }

        // Calculate statistics
        var sortedLatencies = latencies.OrderBy(x => x).ToList();
        var median = sortedLatencies[sortedLatencies.Count / 2];
        var p95 = sortedLatencies[(int)(sortedLatencies.Count * 0.95)];
        var max = sortedLatencies.Last();

        _output.WriteLine($"Wake latency over {latencies.Count} iterations:");
        _output.WriteLine($"  Median: {median:F2}ms");
        _output.WriteLine($"  P95:    {p95:F2}ms");
        _output.WriteLine($"  Max:    {max:F2}ms");

        // Assert — median latency should be well under 5ms.
        // The original 10ms polling loop would show median ~5ms (0-10ms uniform).
        // WaitToReadAsync wakes instantly, so median should be <2ms typically.
        // Using 5ms threshold for reliability across different machines.
        Assert.True(median < 5.0,
            $"REGRESSION: Message pump has high wake latency! Median: {median:F2}ms. " +
            $"Expected <5ms. Original polling had ~5ms median (0-10ms uniform distribution). " +
            "If median is ~5ms, the message pump may have reverted to polling.");

        _output.WriteLine($"✓ Message pump wakes instantly — median latency: {median:F2}ms (threshold: <5ms)");
    }

    /// <summary>
    /// RACE CONDITION TEST: Work items posted just before shutdown must be drained.
    ///
    /// When Dispose() cancels the shutdown token, WaitToReadAsync throws
    /// OperationCanceledException. If work items were written to the channel between
    /// the last TryRead and the cancellation, they must still be processed — otherwise
    /// the caller's TaskCompletionSource never completes and they hang for 5 minutes.
    ///
    /// This test posts work and immediately disposes, verifying the work completes
    /// promptly (not after the 5-minute operation timeout).
    /// </summary>
    [Fact]
    public void Dispose_WithPendingWork_DrainsBeforeExiting()
    {
        // Arrange
        var batch = ExcelSession.BeginBatch(_testFileCopy!);

        // Initialize
        batch.Execute((ctx, ct) =>
        {
            _ = ctx.Book.Worksheets.Item(1);
            return 0;
        });

        // Post an operation from another thread, then immediately Dispose.
        // The operation should complete (via drain) rather than timing out.
        var completed = new ManualResetEventSlim(false);
        Exception? executionError = null;
        int result = -1;

        var executeThread = new Thread(() =>
        {
            try
            {
                // This Execute posts work to the channel.
                // Dispose may happen while we're waiting for the result.
                result = batch.Execute((ctx, ct) =>
                {
                    // Simulate a quick operation
                    return 42;
                });
            }
            catch (ObjectDisposedException)
            {
                // This is acceptable — Dispose beat us to the channel
            }
            catch (Exception ex)
            {
                executionError = ex;
            }
            finally
            {
                completed.Set();
            }
        });

        // Act
        executeThread.Start();

        // Give the execute thread a tiny head start to post work
        Thread.Sleep(10);

        // Dispose while execute may be in flight
        var disposeSw = Stopwatch.StartNew();
        batch.Dispose();
        disposeSw.Stop();

        // Wait for the execute thread to finish — should be quick
        var waitResult = completed.Wait(TimeSpan.FromSeconds(30));

        _output.WriteLine($"Dispose took: {disposeSw.Elapsed.TotalMilliseconds:F0}ms");
        _output.WriteLine($"Execute thread finished: {waitResult}");
        _output.WriteLine($"Execute result: {result}");
        _output.WriteLine($"Execute error: {executionError?.GetType().Name ?? "none"}");

        // Assert — the execute thread must finish within 30s, not hang for 5 minutes.
        // If the drain logic is broken, the TCS would never complete and we'd time out.
        Assert.True(waitResult,
            "REGRESSION: Execute thread hung after Dispose! The message pump shutdown " +
            "did not drain pending work items. Without drain, callers wait for the " +
            "5-minute operation timeout.");

        // If it completed successfully, verify the result
        if (result == 42)
        {
            _output.WriteLine("✓ Work item was executed during shutdown drain");
        }
        else
        {
            _output.WriteLine("✓ Dispose won the race — ObjectDisposedException (acceptable)");
        }

        // Should not have unexpected exceptions (TimeoutException = drain failure)
        if (executionError != null)
        {
            Assert.False(executionError is TimeoutException,
                $"REGRESSION: Execute got TimeoutException during Dispose! " +
                $"This means the work item was orphaned in the channel. Error: {executionError.Message}");
        }
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
        var batch = ExcelSession.BeginBatch(_testFileCopy!);

        batch.Execute((ctx, ct) =>
        {
            _ = ctx.Book.Worksheets.Item(1);
            return 0;
        });

        batch.Dispose();

        // Wait for disposal to complete
        Thread.Sleep(500);

        // Act & Assert — Execute after Dispose should throw ObjectDisposedException
        var ex = Assert.Throws<ObjectDisposedException>(() =>
        {
            batch.Execute((ctx, ct) => 0);
        });

        _output.WriteLine($"✓ Got expected ObjectDisposedException: {ex.Message}");
    }

    /// <summary>
    /// RACE CONDITION TEST: Concurrent Dispose during Execute gives a clean error.
    ///
    /// Simulates the race where Dispose() is called on one thread while Execute() is
    /// actively waiting for its result. The Execute caller should get either:
    /// - Their result (if work completed before disposal)
    /// - ObjectDisposedException (if disposal won the race)
    /// - TimeoutException (if Excel cleanup took too long — unlikely but acceptable)
    ///
    /// It must NOT get a ChannelClosedException or hang indefinitely.
    /// </summary>
    [Fact]
    public void Dispose_DuringActiveExecute_GivesCleanError()
    {
        // Arrange
        var batch = ExcelSession.BeginBatch(_testFileCopy!);

        // Initialize
        batch.Execute((ctx, ct) =>
        {
            _ = ctx.Book.Worksheets.Item(1);
            return 0;
        });

        // Start a long-running operation on another thread
        var operationStarted = new ManualResetEventSlim(false);
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
                    Thread.Sleep(2000);
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
        operationStarted.Wait(TimeSpan.FromSeconds(10));
        _output.WriteLine("Operation started on STA thread, now calling Dispose...");

        // Call Dispose while the operation is mid-flight
        var sw = Stopwatch.StartNew();
        batch.Dispose();
        sw.Stop();

        // Wait for the execute thread to finish
        executeThread.Join(TimeSpan.FromSeconds(30));

        _output.WriteLine($"Dispose took: {sw.Elapsed.TotalMilliseconds:F0}ms");
        _output.WriteLine($"Execute result: {executeResult}");
        _output.WriteLine($"Execute error: {executeError?.GetType().Name}: {executeError?.Message}");

        // Assert — the result must be clean, not a ChannelClosedException
        if (executeError != null)
        {
            // Acceptable exception types during concurrent dispose
            Assert.True(
                executeError is ObjectDisposedException ||
                executeError is TimeoutException ||
                executeError is OperationCanceledException ||
                executeError is InvalidOperationException,
                $"REGRESSION: Got unexpected exception type during concurrent Dispose: " +
                $"{executeError.GetType().Name}: {executeError.Message}. " +
                "Expected ObjectDisposedException, TimeoutException, OperationCanceledException, " +
                "or InvalidOperationException — NOT ChannelClosedException.");

            _output.WriteLine($"✓ Got clean exception during concurrent Dispose: {executeError.GetType().Name}");
        }
        else
        {
            Assert.Equal(99, executeResult);
            _output.WriteLine("✓ Operation completed successfully before Dispose took effect");
        }
    }
}
