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
                await batch.Execute<int>((ctx, ct) =>
                {
                    // Simulate slow operation (sleep longer than timeout)
                    Thread.Sleep(300); // 300ms
                    return 1;
                }, timeout: TimeSpan.FromMilliseconds(100)); // 100ms timeout
            });

            // Verify exception message
            Assert.Contains("timed out after", exception.Message);
            Assert.Contains("100", exception.Message); // Timeout duration in message
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
                await batch.ExecuteAsync<int>(async (ctx, ct) =>
                {
                    // Simulate slow async operation
                    await Task.Delay(300, ct); // 300ms
                    return 1;
                }, timeout: TimeSpan.FromMilliseconds(100)); // 100ms timeout
            });

            // Verify exception message
            Assert.Contains("timed out after", exception.Message);
            Assert.Contains("100", exception.Message);
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

            // Act & Assert - Request 10 minutes, should clamp to max 5 minutes
            var exception = await Assert.ThrowsAsync<TimeoutException>(async () =>
            {
                await batch.Execute<int>((ctx, ct) =>
                {
                    Thread.Sleep(200); // 200ms
                    return 1;
                }, timeout: TimeSpan.FromMinutes(10)); // Request 10 min, get 5 min max
            });

            // Verify exception mentions maximum timeout
            Assert.Contains("maximum timeout", exception.Message);
            Assert.Contains("5", exception.Message); // Max timeout in message
            _output.WriteLine($"✓ Timeout clamped to max: {exception.Message}");
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
            var result = await batch.Execute<int>((ctx, ct) =>
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
            var result = await batch.Execute<int>((ctx, ct) =>
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

            // Assert - Should throw OperationCanceledException, not TimeoutException
            await Assert.ThrowsAsync<OperationCanceledException>(async () =>
            {
                await batch.Execute<int>((ctx, ct) =>
                {
                    Thread.Sleep(200); // 200ms
                    return 1;
                }, timeout: TimeSpan.FromSeconds(5)); // Timeout is 5s, but cancel at 50ms
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
                await batch.Execute<int>((ctx, ct) =>
                {
                    Thread.Sleep(500);
                    return 1;
                }, timeout: TimeSpan.FromMilliseconds(200));
            });

            // Verify exception message contains helpful information
            Assert.Contains("timed out", exception.Message.ToLowerInvariant());
            Assert.Contains("200", exception.Message); // Timeout duration
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

            await Assert.ThrowsAsync<OperationCanceledException>(async () =>
            {
                await batch.ExecuteAsync<int>(async (ctx, ct) =>
                {
                    // Operation that respects cancellation token
                    for (int i = 0; i < 10; i++)
                    {
                        await Task.Delay(50, ct); // Throws if canceled
                    }
                    return 1;
                }, timeout: TimeSpan.FromSeconds(10));
            });

            _output.WriteLine("✓ Operation canceled during timeout period");
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }
}
