using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

/// <summary>
/// Tests for the operation timeout → force-kill → cleanup chain in ExcelBatch.
///
/// These tests validate the fix for Bug 8 (Feb 2026) where a stuck IDispatch.Invoke
/// caused the MCP server to hang permanently because:
/// 1. No timeout recovery existed — ExcelBatch.Dispose() waited forever on STA thread join
/// 2. No pre-emptive kill — Excel process was never killed when operations timed out
/// 3. No session cleanup — WithSessionAsync didn't handle TimeoutException
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test that Execute() throws TimeoutException when operation exceeds timeout
/// - ✅ Test that _operationTimedOut triggers pre-emptive Process.Kill() in Dispose()
/// - ✅ Test that Dispose() completes (doesn't hang) after timeout
/// - ✅ Test that Excel process is cleaned up after timeout + dispose
/// - ✅ Test that cancelled operations also trigger cleanup
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "ExcelBatch")]
[Trait("RunType", "OnDemand")]
[Collection("Sequential")]
public class ExcelBatchTimeoutTests : IAsyncLifetime
{
    private readonly ITestOutputHelper _output;
    private static string? _staticTestFile;
    private string? _testFileCopy;

    public ExcelBatchTimeoutTests(ITestOutputHelper output)
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
                throw new FileNotFoundException($"Static test file not found at {_staticTestFile}.");
            }
        }

        _testFileCopy = Path.Join(Path.GetTempPath(), $"batch-timeout-test-{Guid.NewGuid():N}.xlsx");
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
    /// REGRESSION TEST: Execute() must throw TimeoutException when operation exceeds the configured timeout.
    /// Before Bug 8 fix, timeout existed but had no recovery — the caller got the exception but
    /// Dispose() would then hang forever waiting for the STA thread.
    /// </summary>
    [Fact]
    public void Execute_OperationExceedsTimeout_ThrowsTimeoutException()
    {
        // Arrange — use a very short timeout (3 seconds) to trigger timeout quickly
        var batch = ExcelSession.BeginBatch(
            show: false,
            operationTimeout: TimeSpan.FromSeconds(3),
            _testFileCopy!);

        // Warm up — ensure Excel is ready
        batch.Execute((ctx, ct) =>
        {
            _ = ctx.Book.Worksheets[1];
            return 0;
        });

        _output.WriteLine("Excel initialized, starting long-running operation...");

        // Act & Assert — operation that exceeds timeout must throw TimeoutException
        var sw = Stopwatch.StartNew();
        var ex = Assert.Throws<TimeoutException>(() =>
        {
            batch.Execute((ctx, ct) =>
            {
                // Simulate a hung operation — sleep longer than the timeout
                Thread.Sleep(TimeSpan.FromSeconds(30));
                return 0;
            });
        });
        sw.Stop();

        _output.WriteLine($"TimeoutException thrown after {sw.Elapsed.TotalSeconds:F1}s: {ex.Message}");
        Assert.Contains("timed out", ex.Message, StringComparison.OrdinalIgnoreCase);

        // Dispose must complete and not hang — this is the key regression test
        var disposeSw = Stopwatch.StartNew();
        batch.Dispose();
        disposeSw.Stop();

        _output.WriteLine($"Dispose completed in {disposeSw.Elapsed.TotalSeconds:F1}s");

        // Dispose should complete within a reasonable time (pre-emptive kill + join + wait)
        // Before the fix, Dispose() would hang forever here
        Assert.True(disposeSw.Elapsed < TimeSpan.FromSeconds(30),
            $"REGRESSION: Dispose() took {disposeSw.Elapsed.TotalSeconds:F1}s after timeout — " +
            "pre-emptive kill may not be working. Expected < 30s.");
    }

    /// <summary>
    /// REGRESSION TEST: After timeout, the Excel process must be killed and cleaned up.
    /// Before Bug 8 fix, the hung Excel process would remain alive permanently.
    /// </summary>
    [Fact]
    public void Execute_AfterTimeout_ExcelProcessIsCleaned()
    {
        // Arrange
        var startingProcesses = Process.GetProcessesByName("EXCEL");
        int startingCount = startingProcesses.Length;
        _output.WriteLine($"Excel processes before: {startingCount}");

        var batch = ExcelSession.BeginBatch(
            show: false,
            operationTimeout: TimeSpan.FromSeconds(3),
            _testFileCopy!);

        // Get the Excel process ID before timeout
        int? excelPid = batch.ExcelProcessId;
        _output.WriteLine($"Excel PID for this session: {excelPid}");

        // Warm up
        batch.Execute((ctx, ct) => { _ = ctx.Book.Worksheets[1]; return 0; });

        // Act — trigger timeout
        Assert.Throws<TimeoutException>(() =>
        {
            batch.Execute((ctx, ct) =>
            {
                Thread.Sleep(TimeSpan.FromSeconds(30));
                return 0;
            });
        });

        // Dispose triggers pre-emptive kill
        batch.Dispose();

        // Wait briefly for process cleanup
        Thread.Sleep(2000);

        // Assert — Excel process from this session should be gone
        if (excelPid.HasValue)
        {
            bool processAlive;
            try
            {
                using var process = Process.GetProcessById(excelPid.Value);
                processAlive = !process.HasExited;
            }
            catch (ArgumentException)
            {
                processAlive = false; // Process doesn't exist
            }

            Assert.False(processAlive,
                $"REGRESSION: Excel process {excelPid.Value} is still alive after timeout + dispose. " +
                "Pre-emptive kill in Dispose() may not be working.");

            _output.WriteLine($"✓ Excel process {excelPid.Value} was cleaned up after timeout");
        }

        // Also check total count hasn't leaked
        int endingCount = Process.GetProcessesByName("EXCEL").Length;
        _output.WriteLine($"Excel processes after: {endingCount}");
        Assert.True(endingCount <= startingCount,
            $"Excel process leak! Started with {startingCount}, ended with {endingCount}");
    }

    /// <summary>
    /// REGRESSION TEST: Dispose after timeout must use shorter join timeout (aggressive cleanup).
    /// Before Bug 8 fix, Dispose() used the same 45-second join timeout even when the operation
    /// had already timed out, causing unnecessary delays.
    /// </summary>
    [Fact]
    public void Dispose_AfterTimeout_CompletesWithinAggressiveTimeout()
    {
        // Arrange
        var batch = ExcelSession.BeginBatch(
            show: false,
            operationTimeout: TimeSpan.FromSeconds(3),
            _testFileCopy!);

        batch.Execute((ctx, ct) => { _ = ctx.Book.Worksheets[1]; return 0; });

        // Trigger timeout
        Assert.Throws<TimeoutException>(() =>
        {
            batch.Execute((ctx, ct) =>
            {
                Thread.Sleep(TimeSpan.FromSeconds(30));
                return 0;
            });
        });

        // Act — measure Dispose time
        var sw = Stopwatch.StartNew();
        batch.Dispose();
        sw.Stop();

        _output.WriteLine($"Dispose after timeout completed in {sw.Elapsed.TotalSeconds:F1}s");

        // Assert — with pre-emptive kill + 10s join timeout, Dispose should be fast
        // Before the fix, this could take 45+ seconds or hang forever
        Assert.True(sw.Elapsed < TimeSpan.FromSeconds(25),
            $"REGRESSION: Dispose() took {sw.Elapsed.TotalSeconds:F1}s after timeout. " +
            "Expected < 25s with pre-emptive kill and aggressive 10s join timeout. " +
            "Before Bug 8 fix, this would hang forever.");

        _output.WriteLine("✓ Dispose completed with aggressive timeout (pre-emptive kill working)");
    }

    /// <summary>
    /// REGRESSION TEST: Caller cancellation also triggers aggressive cleanup.
    /// The second catch(OperationCanceledException) in Execute also sets _operationTimedOut.
    /// </summary>
    [Fact]
    public void Execute_CallerCancellation_DisposeCleansUpQuickly()
    {
        // Arrange
        var batch = ExcelSession.BeginBatch(
            show: false,
            operationTimeout: TimeSpan.FromMinutes(5), // Normal timeout — not the trigger
            _testFileCopy!);

        batch.Execute((ctx, ct) => { _ = ctx.Book.Worksheets[1]; return 0; });

        var cts = new CancellationTokenSource();

        // Start a long operation and cancel it after 2 seconds
        var operationStarted = new ManualResetEventSlim(false);
        Exception? caughtException = null;

        var thread = new Thread(() =>
        {
            try
            {
                batch.Execute((ctx, ct) =>
                {
                    operationStarted.Set();
                    // Simulate work that respects cancellation poorly (simulates stuck COM call)
                    Thread.Sleep(TimeSpan.FromSeconds(30));
                    return 0;
                }, cts.Token);
            }
            catch (Exception ex)
            {
                caughtException = ex;
            }
        });

        thread.Start();
        operationStarted.Wait(TimeSpan.FromSeconds(10));

        // Cancel from caller side
        cts.Cancel();
        thread.Join(TimeSpan.FromSeconds(15));

        _output.WriteLine($"Operation exception: {caughtException?.GetType().Name}: {caughtException?.Message}");

        // Act — Dispose should use aggressive cleanup since _operationTimedOut is set
        var sw = Stopwatch.StartNew();
        batch.Dispose();
        sw.Stop();

        _output.WriteLine($"Dispose after cancellation completed in {sw.Elapsed.TotalSeconds:F1}s");

        // Assert — Dispose should not hang
        Assert.True(sw.Elapsed < TimeSpan.FromSeconds(30),
            $"Dispose took {sw.Elapsed.TotalSeconds:F1}s after cancellation — expected < 30s");

        _output.WriteLine("✓ Dispose completed after caller cancellation");
    }

    /// <summary>
    /// REGRESSION TEST: After timeout, subsequent Execute calls must throw TimeoutException
    /// immediately instead of queueing work on the stuck STA thread.
    /// Before this fix, the second caller would queue work and block until its own timeout
    /// expired — causing the entire server to appear hung for up to timeoutSeconds.
    /// </summary>
    [Fact]
    public void Execute_AfterPreviousTimeout_FailsFastWithTimeoutException()
    {
        // Arrange — short timeout to trigger the first timeout quickly
        var batch = ExcelSession.BeginBatch(
            show: false,
            operationTimeout: TimeSpan.FromSeconds(3),
            _testFileCopy!);

        // Warm up
        batch.Execute((ctx, ct) => { _ = ctx.Book.Worksheets[1]; return 0; });

        // Trigger timeout on first operation
        Assert.Throws<TimeoutException>(() =>
        {
            batch.Execute((ctx, ct) =>
            {
                Thread.Sleep(TimeSpan.FromSeconds(30));
                return 0;
            });
        });

        _output.WriteLine("First timeout triggered, now calling Execute again...");

        // Act — second Execute should fail FAST (not wait for its own timeout)
        var sw = Stopwatch.StartNew();
        var ex = Assert.Throws<TimeoutException>(() =>
        {
            batch.Execute((ctx, ct) => { return 42; });
        });
        sw.Stop();

        _output.WriteLine($"Second Execute threw in {sw.Elapsed.TotalMilliseconds:F0}ms: {ex.Message}");

        // Assert — must be near-instant, not another 3+ second timeout wait
        Assert.True(sw.Elapsed < TimeSpan.FromSeconds(1),
            $"REGRESSION: Second Execute took {sw.Elapsed.TotalSeconds:F1}s — expected < 1s. " +
            "The fail-fast pre-check for _operationTimedOut may not be working.");
        Assert.Contains("previous operation", ex.Message, StringComparison.OrdinalIgnoreCase);

        // Cleanup
        batch.Dispose();
        _output.WriteLine("✓ Subsequent Execute after timeout fails fast");
    }

    /// <summary>
    /// Verify that ExcelProcessId is captured during session creation.
    /// This is a prerequisite for the pre-emptive kill to work.
    /// </summary>
    [Fact]
    public void BeginBatch_CapturesExcelProcessId()
    {
        // Arrange & Act
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);

        // Assert
        Assert.NotNull(batch.ExcelProcessId);
        Assert.True(batch.ExcelProcessId > 0, "ExcelProcessId should be a valid PID");

        // Verify the process actually exists
        using var process = Process.GetProcessById(batch.ExcelProcessId.Value);
        Assert.False(process.HasExited, "Excel process should be running");
        Assert.Equal("EXCEL", process.ProcessName, ignoreCase: true);

        _output.WriteLine($"✓ ExcelProcessId captured: {batch.ExcelProcessId}");
    }
}
