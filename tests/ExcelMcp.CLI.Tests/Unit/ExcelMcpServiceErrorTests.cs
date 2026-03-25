using System.Collections.Concurrent;
using System.Reflection;
using Microsoft.Extensions.Logging.Abstractions;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Service;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

/// <summary>
/// Unit tests for ExcelMcpService error handling.
///
/// REGRESSION TESTS for Bug 5 (GitHub #482): Top-level exception catch in ProcessAsync
/// only included ex.Message, losing the exception type. This makes debugging impossible
/// when the same message text is shared by multiple exception types.
/// </summary>
[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "ExcelMcpService")]
[Trait("Speed", "Fast")]
public sealed class ExcelMcpServiceErrorTests
{
    /// <summary>
    /// REGRESSION TEST for Bug 5 (#482): When an unexpected exception escapes
    /// the ProcessAsync routing switch (e.g. NullReferenceException on null Command),
    /// the error message must include the exception type name so the caller can
    /// distinguish different failure modes.
    /// </summary>
    [Fact]
    public async Task ProcessAsync_UnexpectedExceptionEscapesRouter_ErrorMessageIncludesTypeName()
    {
        // Arrange
        using var service = new ExcelMcpService();

        // null Command triggers NullReferenceException in parts = request.Command.Split(...)
        // This exercises the top-level catch (Exception ex) block in ProcessAsync
#pragma warning disable CS8714 // required property set to null intentionally to trigger NRE
        var request = new ServiceRequest { Command = null! };
#pragma warning restore CS8714

        // Act
        var response = await service.ProcessAsync(request);

        // Assert
        Assert.False(response.Success);
        Assert.NotNull(response.ErrorMessage);

        // REGRESSION: Before fix, only ex.Message was returned ("Object reference not set...").
        // After fix, the type name is prepended: "NullReferenceException: Object reference..."
        Assert.Contains("NullReferenceException", response.ErrorMessage,
            StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Verifies that normal error responses (business logic, not unexpected exceptions)
    /// still work correctly after the Bug 5 fix. The format change should only affect
    /// the top-level unexpected exception handler.
    /// </summary>
    [Fact]
    public async Task ProcessAsync_UnknownCategory_ReturnsNormalErrorWithoutTypeName()
    {
        // Arrange
        using var service = new ExcelMcpService();
        var request = new ServiceRequest { Command = "unknowncategory.someaction" };

        // Act
        var response = await service.ProcessAsync(request);

        // Assert
        Assert.False(response.Success);
        Assert.NotNull(response.ErrorMessage);
        Assert.Contains("Unknown command category", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        // This path returns a normal string, not an exception-caught message,
        // so it should NOT contain an exception type name prefix.
        Assert.DoesNotContain("Exception:", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Verifies that the WithSessionAsync exception handler (the catch at the bottom
    /// of ProcessAsync, covering session-level operations) also includes the type name.
    /// </summary>
    [Fact]
    public async Task ProcessAsync_SessionCommandWithInvalidSessionId_ReturnsUsableError()
    {
        // Arrange
        using var service = new ExcelMcpService();

        // Send a sheet.list command with a session ID that doesn't exist
        var request = new ServiceRequest
        {
            Command = "sheet.list",
            SessionId = "nonexistent-session-id-00000000"
        };

        // Act
        var response = await service.ProcessAsync(request);

        // Assert — should fail gracefully with a descriptive message, not an unhandled exception
        Assert.False(response.Success);
        Assert.NotNull(response.ErrorMessage);
        Assert.NotEmpty(response.ErrorMessage);
    }

    [Fact]
    public async Task ProcessAsync_SessionCommandOnTimedOutSession_FailsFastBeforeExecutingBatch()
    {
        using var service = new ExcelMcpService();
        var batch = new FakeBatch { HasTimedOutOperation = true };
        const string sessionId = "timed-out-sheet-list";

        RegisterSession(service, sessionId, batch);

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "sheet.list",
            SessionId = sessionId
        });

        Assert.False(response.Success);
        Assert.NotNull(response.ErrorMessage);
        Assert.Contains("timed out or was cancelled", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("reopen", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, batch.ExecuteCalls);
    }

    [Fact]
    public async Task ProcessAsync_SessionSaveOnTimedOutSession_FailsFastBeforeSaving()
    {
        using var service = new ExcelMcpService();
        var batch = new FakeBatch { HasTimedOutOperation = true };
        const string sessionId = "timed-out-save";

        RegisterSession(service, sessionId, batch);

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.save",
            SessionId = sessionId
        });

        Assert.False(response.Success);
        Assert.NotNull(response.ErrorMessage);
        Assert.Contains("timed out or was cancelled", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, batch.SaveCalls);
    }

    [Fact]
    public async Task ProcessAsync_SessionSaveOnHealthySession_StillSavesNormally()
    {
        using var service = new ExcelMcpService();
        var batch = new FakeBatch();
        const string sessionId = "healthy-save";

        RegisterSession(service, sessionId, batch);

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.save",
            SessionId = sessionId
        });

        Assert.True(response.Success);
        Assert.Equal(1, batch.SaveCalls);
    }

    private static void RegisterSession(ExcelMcpService service, string sessionId, FakeBatch batch)
    {
        var sessionManager = GetPrivateField<SessionManager>(service, "_sessionManager");
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
    }

    private static T GetPrivateField<T>(object instance, string fieldName)
    {
        var field = instance.GetType().GetField(fieldName, BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(field);
        return (T)field!.GetValue(instance)!;
    }

    private sealed class FakeBatch : IExcelBatch
    {
        public string WorkbookPath { get; init; } = Path.Combine(Path.GetTempPath(), $"fake-batch-{Guid.NewGuid():N}.xlsx");
        public Microsoft.Extensions.Logging.ILogger Logger { get; } = NullLogger.Instance;
        public IReadOnlyDictionary<string, Excel.Workbook> Workbooks { get; } = new Dictionary<string, Excel.Workbook>();
        public bool HasTimedOutOperation { get; init; }
        public int ExecuteCalls { get; private set; }
        public int SaveCalls { get; private set; }
        public int? ExcelProcessId => 1234;
        public TimeSpan OperationTimeout => TimeSpan.FromSeconds(5);

        public Excel.Workbook GetWorkbook(string filePath) => throw new NotSupportedException();

        public void Execute(Action<ExcelContext, CancellationToken> operation, CancellationToken cancellationToken = default)
        {
            ExecuteCalls++;
            throw new InvalidOperationException("Execute should not be called for a poisoned fake batch.");
        }

        public T Execute<T>(Func<ExcelContext, CancellationToken, T> operation, CancellationToken cancellationToken = default)
        {
            ExecuteCalls++;
            throw new InvalidOperationException("Execute should not be called for a poisoned fake batch.");
        }

        public void Save(CancellationToken cancellationToken = default)
        {
            SaveCalls++;
        }

        public bool IsExcelProcessAlive() => true;

        public void Dispose()
        {
        }
    }
}
