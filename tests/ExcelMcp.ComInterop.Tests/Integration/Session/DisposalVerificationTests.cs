using System.Collections.Concurrent;
using Microsoft.Extensions.Logging;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration;

/// <summary>
/// Verifies that the Interlocked disposal fix prevents double disposal.
/// This test uses a custom logger to capture and verify disposal messages.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "SessionManager")]
[Trait("RequiresExcel", "true")]
[Collection("Sequential")]
public class DisposalVerificationTest : IAsyncLifetime
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly List<string> _testFiles = new();

    public DisposalVerificationTest(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Combine(Path.GetTempPath(), $"DisposalVerificationTest_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public Task InitializeAsync() => Task.CompletedTask;

    public Task DisposeAsync()
    {
        foreach (var file in _testFiles)
        {
            File.Delete(file);
        }

        Directory.Delete(_tempDir);

        return Task.CompletedTask;
    }

    /// <summary>
    /// Path to the template xlsx file used for fast test file creation.
    /// Copying a template is ~1000x faster than spawning Excel to create a new workbook.
    /// </summary>
    private static readonly string TemplateFilePath = Path.Combine(
        Path.GetDirectoryName(typeof(DisposalVerificationTest).Assembly.Location)!,
        "Integration", "Session", "TestFiles", "batch-test-static.xlsx");

    private string CreateTestFile(string testName)
    {
        var fileName = $"{testName}_{Guid.NewGuid():N}.xlsx";
        var filePath = Path.Combine(_tempDir, fileName);

        // PERFORMANCE OPTIMIZATION: Copy from template instead of spawning Excel.
        // This reduces test file creation from ~7-14 seconds to <10ms.
        File.Copy(TemplateFilePath, filePath);

        _testFiles.Add(filePath);
        return filePath;
    }

    [Fact]
    public void Dispose_CalledTwice_OnlyDisposesOnce()
    {
        var testFile = CreateTestFile(nameof(Dispose_CalledTwice_OnlyDisposesOnce));

        // Create logger that captures messages
        using var owned = new OwnedExcelProcessScope();
        var messages = new ConcurrentQueue<string>();
        using var loggerFactory = LoggerFactory.Create(builder =>
        {
            builder.AddProvider(new TestLoggerProvider(_output, messages));
            builder.SetMinimumLevel(LogLevel.Debug);
        });
        var logger = loggerFactory.CreateLogger<ExcelBatch>();

        // Create batch with logger
        using var batch = new ExcelBatch(new[] { testFile }, logger);
        Assert.Equal(testFile, batch.Execute((context, _) => context.Book.FullName));

        // First disposal - should execute
        _output.WriteLine("=== First DisposeAsync call ===");
        batch.Dispose();
        _output.WriteLine("=== First DisposeAsync completed ===");

        // Second disposal - should be no-op (return immediately)
        _output.WriteLine("=== Second DisposeAsync call ===");
        batch.Dispose();
        _output.WriteLine("=== Second DisposeAsync completed ===");

        // Third disposal - should also be no-op
        _output.WriteLine("=== Third DisposeAsync call ===");
        batch.Dispose();
        _output.WriteLine("=== Third DisposeAsync completed ===");

        Assert.Single(messages, message => message.Contains("Dispose starting for", StringComparison.Ordinal));
        Assert.Equal(2, messages.Count(message => message.Contains("Dispose skipped - already disposed", StringComparison.Ordinal)));
        Assert.Single(messages, message => message.Contains("STA thread cleanup completed for", StringComparison.Ordinal));
        owned.AssertAllExited();
    }

    [Fact]
    public void SessionManager_DoubleDisposal_OnlyDisposesOnce()
    {
        var testFile = CreateTestFile(nameof(SessionManager_DoubleDisposal_OnlyDisposesOnce));

        // This mimics the original bug scenario:
        // 1. User calls CloseSessionAsync (triggers batch.DisposeAsync)
        // 2. using manager disposes

        using var owned = new OwnedExcelProcessScope();
        using var manager = new SessionManager();

        _output.WriteLine("Creating session...");
        var sessionId = manager.CreateSession(testFile);
        var batch = Assert.IsAssignableFrom<IExcelBatch>(manager.GetSession(sessionId));
        Assert.Equal(testFile, batch.Execute((context, _) => context.Book.FullName));
        _output.WriteLine($"Session created: {sessionId}");

        // This calls batch.DisposeAsync internally
        _output.WriteLine("Calling CloseSession (first disposal)...");
        Assert.True(manager.CloseSession(sessionId));
        Assert.Equal(0, manager.ActiveSessionCount);
        _output.WriteLine("CloseSession completed");

        // await using will call manager.DisposeAsync at end of scope
        // Since we already removed the batch from the dictionary in CloseSessionAsync,
        // the batch won't be disposed again
        _output.WriteLine("Exiting using scope (manager disposal)...");
        manager.Dispose();
        Assert.Equal(0, manager.ActiveSessionCount);
        owned.AssertAllExited();
    }
}

/// <summary>
/// Custom logger provider that writes to xUnit output
/// </summary>
internal sealed class TestLoggerProvider : ILoggerProvider
{
    private readonly ITestOutputHelper _output;
    private readonly ConcurrentQueue<string> _messages;

    public TestLoggerProvider(ITestOutputHelper output, ConcurrentQueue<string> messages)
    {
        _output = output;
        _messages = messages;
    }

    public ILogger CreateLogger(string categoryName)
    {
        return new TestLogger(_output, categoryName, _messages);
    }

    public void Dispose()
    {
    }
}

/// <summary>
/// Custom logger that writes to xUnit output
/// </summary>
internal sealed class TestLogger : ILogger
{
    private readonly ITestOutputHelper _output;
    private readonly string _categoryName;
    private readonly ConcurrentQueue<string> _messages;

    public TestLogger(ITestOutputHelper output, string categoryName, ConcurrentQueue<string> messages)
    {
        _output = output;
        _categoryName = categoryName;
        _messages = messages;
    }

    public IDisposable? BeginScope<TState>(TState state) where TState : notnull
    {
        return null;
    }

    public bool IsEnabled(LogLevel logLevel)
    {
        return true;
    }

    public void Log<TState>(
        LogLevel logLevel,
        EventId eventId,
        TState state,
        Exception? exception,
        Func<TState, Exception?, string> formatter)
    {
        var message = formatter(state, exception);
        _messages.Enqueue(message);
        _output.WriteLine($"[{logLevel}] {_categoryName}: {message}");
    }
}


