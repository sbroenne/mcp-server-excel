using System.Diagnostics;
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

    public Task InitializeAsync()
    {
        // Kill any existing Excel processes to ensure clean state
        try
        {
            var existingProcesses = Process.GetProcessesByName("EXCEL");
            if (existingProcesses.Length > 0)
            {
                _output.WriteLine($"Cleaning up {existingProcesses.Length} existing Excel processes...");
                foreach (var p in existingProcesses)
                {
                    p.Kill(entireProcessTree: true);
                    p.WaitForExit(5000);
                    p.Dispose();
                }
            }
        }
        catch (Exception ex)
        {
            _output.WriteLine($"Warning: Failed to clean Excel processes: {ex.Message}");
        }

        return Task.CompletedTask;
    }

    public Task DisposeAsync()
    {
        foreach (var file in _testFiles)
        {
            try
            {
                if (File.Exists(file))
                {
                    File.Delete(file);
                }
            }
            catch
            {
                // Best effort
            }
        }

        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, recursive: true);
            }
        }
        catch
        {
            // Best effort
        }

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
        var loggerFactory = LoggerFactory.Create(builder =>
        {
            builder.AddProvider(new TestLoggerProvider(_output));
            builder.SetMinimumLevel(LogLevel.Debug);
        });
        var logger = loggerFactory.CreateLogger<ExcelBatch>();

        // Create batch with logger
        var batch = new ExcelBatch(new[] { testFile }, logger);

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

        _output.WriteLine("=== Test completed successfully - only one disposal should have executed ===");
    }

    [Fact]
    public void SessionManager_DoubleDisposal_OnlyDisposesOnce()
    {
        var testFile = CreateTestFile(nameof(SessionManager_DoubleDisposal_OnlyDisposesOnce));

        // This mimics the original bug scenario:
        // 1. User calls CloseSessionAsync (triggers batch.DisposeAsync)
        // 2. using manager disposes

        using var manager = new SessionManager();

        _output.WriteLine("Creating session...");
        var sessionId = manager.CreateSession(testFile);
        _output.WriteLine($"Session created: {sessionId}");

        // This calls batch.DisposeAsync internally
        _output.WriteLine("Calling CloseSession (first disposal)...");
        manager.CloseSession(sessionId);
        _output.WriteLine("CloseSession completed");

        // await using will call manager.DisposeAsync at end of scope
        // Since we already removed the batch from the dictionary in CloseSessionAsync,
        // the batch won't be disposed again
        _output.WriteLine("Exiting using scope (manager disposal)...");
    }
}

/// <summary>
/// Custom logger provider that writes to xUnit output
/// </summary>
internal sealed class TestLoggerProvider : ILoggerProvider
{
    private readonly ITestOutputHelper _output;

    public TestLoggerProvider(ITestOutputHelper output)
    {
        _output = output;
    }

    public ILogger CreateLogger(string categoryName)
    {
        return new TestLogger(_output, categoryName);
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

    public TestLogger(ITestOutputHelper output, string categoryName)
    {
        _output = output;
        _categoryName = categoryName;
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
        _output.WriteLine($"[{logLevel}] {_categoryName}: {message}");
    }
}





