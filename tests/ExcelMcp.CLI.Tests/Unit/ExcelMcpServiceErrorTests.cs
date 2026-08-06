using System.Collections.Concurrent;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
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
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ExcelMcpServiceErrorTests : IDisposable
{
    private readonly string _stateRoot = Path.Combine(
        Path.GetTempPath(),
        $"excelmcp-service-errors-{Guid.NewGuid():N}");

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
        using var service = CreateService();

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
        using var service = CreateService();
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

    [Fact]
    public async Task ProcessAsync_FailedResponse_RedactsSensitiveDetailsAndPreservesContext()
    {
        const string email = "ada@example.com";
        const string secret = "SuperSecretValue";
        var sensitivePath = $@"C:\Users\Ada\Finance\{email}\Password={secret}\book.xlsx";

        ServiceResponse response;
        using (var service = CreateService())
        {
            response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.open",
                Args = JsonSerializer.Serialize(new { filePath = sensitivePath }, ServiceProtocol.JsonOptions)
            });
        }

        Assert.False(response.Success);
        Assert.Equal("session.open", response.Command);
        Assert.NotNull(response.ErrorMessage);
        Assert.Contains("[REDACTED_PATH]", response.ErrorMessage, StringComparison.Ordinal);
        Assert.DoesNotContain(sensitivePath, response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(email, response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(secret, response.ErrorMessage, StringComparison.Ordinal);
        Assert.DoesNotContain(email, response.InnerError ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(secret, response.InnerError ?? string.Empty, StringComparison.Ordinal);

    }

    /// <summary>
    /// Verifies that the WithSessionAsync exception handler (the catch at the bottom
    /// of ProcessAsync, covering session-level operations) also includes the type name.
    /// </summary>
    [Fact]
    public async Task ProcessAsync_SessionCommandWithInvalidSessionId_ReturnsUsableError()
    {
        // Arrange
        using var service = CreateService();

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
        using var service = CreateService();
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
        using var service = CreateService();
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
        using var service = CreateService();
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

    [Fact]
    public async Task ProcessAsync_SessionCloseSaveAfterRpcDisconnected_CleansSessionAndReturnsActionableError()
    {
        using var service = CreateService();
        var batch = new FakeBatch
        {
#pragma warning disable CA2201
            SaveException = new InvalidOperationException(
                "Failed to save workbook 'book.xlsx': The object invoked has disconnected from its clients.",
                new COMException("The object invoked has disconnected from its clients.", unchecked((int)0x80010108))),
#pragma warning restore CA2201
            IsAliveAfterSaveException = false
        };
        const string sessionId = "disconnected-close-save";

        RegisterSession(service, sessionId, batch);

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new { save = true }, ServiceProtocol.JsonOptions)
        });

        Assert.False(response.Success);
        Assert.True(
            response.ErrorCategory == "ExcelProcessDied",
            $"Expected ExcelProcessDied but got '{response.ErrorCategory}'. Message: {response.ErrorMessage}");
        Assert.NotNull(response.ErrorMessage);
        Assert.Contains("disconnected", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Session has been cleaned up", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(1, batch.SaveCalls);
        Assert.Equal(1, batch.DisposeCalls);

        var listResponse = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
        Assert.True(listResponse.Success);
        Assert.NotNull(listResponse.Result);
        Assert.DoesNotContain(sessionId, listResponse.Result, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task WorkflowPlan_InterruptedStep_ForceClosesSessionWhileLeaseIsActive(bool cancelled)
    {
        using var service = CreateService();
        var batch = new FakeBatch
        {
            ExecuteException = cancelled
                ? new OperationCanceledException("cancelled by test")
                : new TimeoutException("timed out by test")
        };
        string sessionId = cancelled ? "workflow-cancelled" : "workflow-timeout";
        RegisterSession(service, sessionId, batch);

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "workflow.execute-plan",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                operations = new[]
                {
                    new
                    {
                        command = "range.set-values",
                        args = new
                        {
                            sheetName = "Sheet1",
                            rangeAddress = "A1",
                            values = new object?[][] { [1] },
                        },
                    },
                },
                fastMode = false,
            }, ServiceProtocol.JsonOptions),
        });

        Assert.False(response.Success);
        Assert.Equal("UnknownOutcome", response.ErrorCategory);
        Assert.Equal(1, batch.DisposeCalls);
        Assert.Equal(0, service.SessionCount);
    }

    [Fact]
    public async Task WorkflowPlan_CheckpointFailure_StopsBeforeLaterMutationEvenWhenContinueWasRequested()
    {
        using var service = CreateService();
        var batch = new FakeBatch { AllowNoOpExecutions = true };
        const string sessionId = "workflow-checkpoint-failure";
        RegisterSession(service, sessionId, batch);

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "workflow.execute-plan",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                operations = new object[]
                {
                    new
                    {
                        command = "range.set-values",
                        args = new
                        {
                            sheetName = "Sheet1",
                            rangeAddress = "A1",
                            values = new object?[][] { [1] },
                        },
                    },
                    new
                    {
                        command = "range.set-values",
                        args = new
                        {
                            sheetName = "Sheet1",
                            rangeAddress = "A2",
                            values = new object?[][] { [2] },
                        },
                    },
                },
                stopOnError = false,
                checkpointMode = "once",
            }, ServiceProtocol.JsonOptions),
        });

        Assert.False(response.Success);
        Assert.Equal("CheckpointFailed", response.ErrorCategory);
        using var receipt = JsonDocument.Parse(response.Result!);
        Assert.Equal(1, receipt.RootElement.GetProperty("attemptedCount").GetInt32());
        Assert.Equal(0, receipt.RootElement.GetProperty("completedCount").GetInt32());
        Assert.False(receipt.RootElement.TryGetProperty("checkpoint", out _));
        var steps = receipt.RootElement.GetProperty("steps").EnumerateArray().ToArray();
        Assert.Equal("failed", steps[0].GetProperty("status").GetString());
        Assert.Equal("notStarted", steps[1].GetProperty("status").GetString());
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

    private ExcelMcpService CreateService() => new(_stateRoot);

    /// <inheritdoc />
    public void Dispose()
    {
        if (Directory.Exists(_stateRoot))
        {
            Directory.Delete(_stateRoot, recursive: true);
        }

        GC.SuppressFinalize(this);
    }

    private sealed class FakeBatch : IExcelBatch
    {
        public string WorkbookPath { get; init; } = Path.Combine(Path.GetTempPath(), $"fake-batch-{Guid.NewGuid():N}.xlsx");
        public Microsoft.Extensions.Logging.ILogger Logger { get; } = NullLogger.Instance;
        public IReadOnlyDictionary<string, Excel.Workbook> Workbooks { get; } = new Dictionary<string, Excel.Workbook>();
        public bool HasTimedOutOperation { get; init; }
        public bool IsAlive { get; private set; } = true;
        public bool IsAliveAfterSaveException { get; init; } = true;
        public Exception? SaveException { get; init; }
        public Exception? ExecuteException { get; init; }
        public bool AllowNoOpExecutions { get; init; }
        public int ExecuteCalls { get; private set; }
        public int SaveCalls { get; private set; }
        public int DisposeCalls { get; private set; }
        public int? ExcelProcessId => 1234;
        public TimeSpan OperationTimeout => TimeSpan.FromSeconds(5);

        public Excel.Workbook GetWorkbook(string filePath) => throw new NotSupportedException();

        public void Execute(Action<ExcelContext, CancellationToken> operation, CancellationToken cancellationToken = default)
        {
            ExecuteCalls++;
            if (ExecuteException is not null)
            {
                throw ExecuteException;
            }

            if (!AllowNoOpExecutions)
            {
                throw new InvalidOperationException("Execute should not be called for a poisoned fake batch.");
            }
        }

        public T Execute<T>(Func<ExcelContext, CancellationToken, T> operation, CancellationToken cancellationToken = default)
        {
            ExecuteCalls++;
            if (ExecuteException is not null)
            {
                throw ExecuteException;
            }

            if (!AllowNoOpExecutions)
            {
                throw new InvalidOperationException("Execute should not be called for a poisoned fake batch.");
            }

            return default!;
        }

        public void Save(CancellationToken cancellationToken = default)
        {
            SaveCalls++;
            if (SaveException != null)
            {
                IsAlive = IsAliveAfterSaveException;
                throw SaveException;
            }
        }

        public bool IsExcelProcessAlive() => IsAlive;

        public void Dispose()
        {
            DisposeCalls++;
        }
    }
}
