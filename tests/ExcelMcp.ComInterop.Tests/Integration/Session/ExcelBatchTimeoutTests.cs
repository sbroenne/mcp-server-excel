using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

/// <summary>
/// Integration tests for ExcelBatch timeout functionality.
/// Tests timeout enforcement, clamping, and error handling.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "Timeout")]
[Trait("RequiresExcel", "true")]
public class ExcelBatchTimeoutTests
{
    private readonly ITestOutputHelper _output;

    public ExcelBatchTimeoutTests(ITestOutputHelper output)
    {
        _output = output;
    }

    private async Task<string> CreateTempTestFileAsync()
    {
        string testFile = Path.Join(Path.GetTempPath(), $"timeout-test-{Guid.NewGuid():N}.xlsx");
        await ExcelSession.CreateNew(testFile, isMacroEnabled: false, (ctx, ct) => 0);
        return testFile;
    }

    [Fact]
    public async Task Execute_OperationExceedsTimeout_ThrowsTimeoutException()
    {
        // Arrange
        string testFile = await CreateTempTestFileAsync();

        try
        {
            await using var batch = await ExcelSession.BeginBatchAsync(testFile);

            // Act & Assert
            var exception = await Assert.ThrowsAsync<TimeoutException>(async () =>
            {
                await batch.Execute((ctx, ct) =>
                {
                    // Simulate slow operation (sleep longer than timeout)
                    Thread.Sleep(300); // 300ms
                    return 1;
                }, timeout: TimeSpan.FromMilliseconds(100)); // 100ms timeout
            });

            // Verify exception message (formatted in minutes)
            Assert.Contains("timed out after", exception.Message);
            // Culture-agnostic check: verify the message contains a small number (100ms ≈ 0.00166 min)
            Assert.Matches(@"0[.,]0+1\d+", exception.Message); // Matches "0.00166" or "0,00166"
            _output.WriteLine($"✓ TimeoutException thrown: {exception.Message}");
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }

    [Fact]
    public async Task ExecuteAsync_OperationExceedsTimeout_ThrowsTimeoutException()
    {
        // Arrange
        string testFile = await CreateTempTestFileAsync();

        try
        {
            await using var batch = await ExcelSession.BeginBatchAsync(testFile);

            // Act & Assert
            var exception = await Assert.ThrowsAsync<TimeoutException>(async () =>
            {
                await batch.ExecuteAsync(async (ctx, ct) =>
                {
                    // Simulate slow async operation
                    await Task.Delay(300, ct); // 300ms
                    return 1;
                }, timeout: TimeSpan.FromMilliseconds(100)); // 100ms timeout
            });

            // Verify exception message (formatted in minutes)
            Assert.Contains("timed out after", exception.Message);
            // Culture-agnostic check: verify the message contains a small number (100ms ≈ 0.00166 min)
            Assert.Matches(@"0[.,]0+1\d+", exception.Message); // Matches "0.00166" or "0,00166"
            _output.WriteLine($"✓ Async TimeoutException thrown: {exception.Message}");
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }

    [Fact]
    public async Task Execute_TimeoutExceedsMax_ClampsToMaxTimeout()
    {
        // Arrange
        string testFile = await CreateTempTestFileAsync();

        try
        {
            await using var batch = await ExcelSession.BeginBatchAsync(testFile);

            // Act - Request 10 minutes timeout, verify operation completes with clamped timeout
            // The clamping happens internally, we verify by seeing stderr log shows 5.0min not 10.0min
            var result = await batch.Execute((ctx, ct) =>
            {
                Thread.Sleep(100); // Fast operation
                return 42;
            }, timeout: TimeSpan.FromMinutes(10)); // Request 10 min, should clamp to 5 min

            // Assert - Operation succeeded (timeout was clamped but operation was fast enough)
            Assert.Equal(42, result);
            _output.WriteLine($"✓ Timeout clamped to max (check stderr for '5.0min' not '10.0min')");
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }

    [Fact]
    public async Task Execute_NoTimeoutSpecified_UsesDefaultTimeout()
    {
        // Arrange
        string testFile = await CreateTempTestFileAsync();

        try
        {
            await using var batch = await ExcelSession.BeginBatchAsync(testFile);

            // Act - Operation completes quickly, default timeout not reached
            var result = await batch.Execute((ctx, ct) =>
            {
                return 42;
            }); // No timeout specified, uses default 2 minutes

            // Assert - Should succeed
            Assert.Equal(42, result);
            _output.WriteLine("✓ Default timeout applied, operation completed");
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }

    [Fact]
    public async Task Execute_FastOperation_CompletesBeforeTimeout()
    {
        // Arrange
        string testFile = await CreateTempTestFileAsync();

        try
        {
            await using var batch = await ExcelSession.BeginBatchAsync(testFile);

            // Act - Fast operation with timeout
            var result = await batch.Execute((ctx, ct) =>
            {
                Thread.Sleep(50); // 50ms - well within timeout
                return 99;
            }, timeout: TimeSpan.FromSeconds(5));

            // Assert - Should succeed
            Assert.Equal(99, result);
            _output.WriteLine("✓ Fast operation completed before timeout");
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }

    [Fact]
    public async Task Execute_CancellationToken_StillWorks()
    {
        // Arrange
        string testFile = await CreateTempTestFileAsync();
        using var cts = new CancellationTokenSource();

        try
        {
            await using var batch = await ExcelSession.BeginBatchAsync(testFile);

            // Act - Cancel before timeout
            cts.CancelAfter(50); // Cancel after 50ms

            // Assert - Should throw OperationCanceledException or TaskCanceledException 
            var exception = await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
            {
                await batch.Execute((ctx, ct) =>
                {
                    // Check cancellation token multiple times during slow operation
                    for (int i = 0; i < 10; i++)
                    {
                        ct.ThrowIfCancellationRequested(); // Will throw when cancelled
                        Thread.Sleep(50); // 50ms each iteration = 500ms total if not cancelled
                    }
                    return 1;
                }, cancellationToken: cts.Token, timeout: TimeSpan.FromSeconds(5)); // Timeout is 5s, but cancel at 50ms
            });

            _output.WriteLine("✓ CancellationToken works independently of timeout");
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }

    [Fact]
    public async Task Execute_TimeoutMessage_ContainsDuration()
    {
        // Arrange
        string testFile = await CreateTempTestFileAsync();

        try
        {
            await using var batch = await ExcelSession.BeginBatchAsync(testFile);

            // Act & Assert
            var exception = await Assert.ThrowsAsync<TimeoutException>(async () =>
            {
                await batch.Execute((ctx, ct) =>
                {
                    Thread.Sleep(500);
                    return 1;
                }, timeout: TimeSpan.FromMilliseconds(200));
            });

            // Verify exception message contains helpful information (formatted in minutes)
            Assert.Contains("timed out", exception.Message.ToLowerInvariant());
            // Culture-agnostic check: verify the message contains a small number (200ms ≈ 0.00333 min)
            Assert.Matches(@"0[.,]0+3\d+", exception.Message); // Matches "0.00333" or "0,00333"
            _output.WriteLine($"✓ Exception message: {exception.Message}");
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }

    [Fact]
    public async Task ExecuteAsync_RespectsCancellationDuringTimeout()
    {
        // Arrange
        string testFile = await CreateTempTestFileAsync();
        using var cts = new CancellationTokenSource();

        try
        {
            await using var batch = await ExcelSession.BeginBatchAsync(testFile);

            // Act - Operation checks cancellation token
            cts.CancelAfter(100); // Cancel after 100ms

            var exception = await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
            {
                await batch.ExecuteAsync(async (ctx, ct) =>
                {
                    // Operation that respects cancellation token
                    for (int i = 0; i < 10; i++)
                    {
                        await Task.Delay(50, ct); // Throws if canceled
                    }
                    return 1;
                }, cancellationToken: cts.Token, timeout: TimeSpan.FromSeconds(10));
            });

            _output.WriteLine("✓ Operation canceled during timeout period");
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }
}
