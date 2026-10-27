using System.Collections.Concurrent;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.Json;
using Microsoft.Extensions.Logging.Abstractions;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Utilities;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

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
[Trait("RequiresExcel", "false")]
public sealed class ExcelMcpServiceErrorTests
{
    [Fact]
    public void SessionRetention_DefersLateSaveAsPathReconciliationDuringShutdown()
    {
        using var service = new ExcelMcpService();
        var manager = service.SessionManager;
        var batch = (ExcelBatch)RuntimeHelpers.GetUninitializedObject(typeof(ExcelBatch));
        var original = Path.GetFullPath("retained-save-original.xlsx");
        var target = Path.GetFullPath("retained-save-target.xlsx");
        foreach (var (name, value) in new (string, object)[]
        {
            ("_workbookPath", original),
            ("_startupCompletion", Task.CompletedTask),
            ("_statusLock", new object()),
            ("_operationState", "completed"),
            ("_waitExpired", true)
        })
        {
            var field = typeof(ExcelBatch).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.NotNull(field);
            field.SetValue(batch, value);
        }
        const string session = "late-save-as";
        RegisterSession(service, session, batch);
        manager.ReserveSessionFilePath(session, target);
        manager.RetainPendingSaveAs(session);
        typeof(ExcelBatch).GetField("_workbookPath", BindingFlags.Instance | BindingFlags.NonPublic)!
            .SetValue(batch, target);
        var disposed = typeof(SessionManager).GetField("_disposed", BindingFlags.Instance | BindingFlags.NonPublic)!;
        try
        {
            disposed.SetValue(manager, true);
            Assert.True(manager.IsSessionRetained(session));
            Assert.True(manager.IsSessionRetained(session));
            Assert.Equal(original, GetPrivateField<ConcurrentDictionary<string, string>>(
                manager, "_sessionFilePaths")[session]);
            Assert.Equal(target, GetPrivateField<ConcurrentDictionary<string, string>>(
                manager, "_sessionFilePathReservations")[session]);

            disposed.SetValue(manager, false);
            Assert.True(manager.IsSessionRetained(session));
            Assert.True(manager.TryGetFilePath(session, out var reconciled));
            Assert.Equal(target, reconciled);
            Assert.Empty(GetPrivateField<ConcurrentDictionary<string, string>>(manager, "_sessionFilePathReservations"));
            var claims = GetPrivateField<ConcurrentDictionary<string, string>>(manager, "_activeFilePaths");
            Assert.False(claims.ContainsKey(original));
            Assert.Equal(session, claims[target]);
        }
        finally
        {
            disposed.SetValue(manager, false);
            GetPrivateField<ConcurrentDictionary<string, IExcelBatch>>(manager, "_activeSessions")
                .TryRemove(session, out _);
        }
    }

    [Fact]
    public void SessionRetention_PreservesCompletedLateStartupFailure()
    {
        using var service = new ExcelMcpService();
        var batch = (ExcelBatch)RuntimeHelpers.GetUninitializedObject(typeof(ExcelBatch));
        const string message = "Opening failed after the caller stopped waiting.";
        foreach (var (name, value) in new (string, object)[]
        {
            ("_workbookPath", Path.GetFullPath("late-startup-failure.xlsx")),
            ("_excelProcessId", int.MaxValue),
            ("_startupCompletion", Task.FromException(new InvalidOperationException(message))),
            ("_statusLock", new object()),
            ("_operationState", "failed"),
            ("_operationError", message),
            ("_waitExpired", true)
        })
        {
            var field = typeof(ExcelBatch).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.NotNull(field);
            field.SetValue(batch, value);
        }
        const string sessionId = "late-startup-failure";
        RegisterSession(service, sessionId, batch);
        try
        {
            Assert.False(batch.IsExcelProcessAlive());
            Assert.True(service.SessionManager.IsSessionRetained(sessionId));
            Assert.True(service.SessionManager.IsSessionRetained(sessionId));
            Assert.False(service.SessionManager.TryBeginOperation(sessionId, out _, out var error));
            Assert.Contains(message, error, StringComparison.Ordinal);
            var status = Assert.IsType<ExcelOperationStatus>(service.SessionManager.GetOperationStatus(sessionId));
            Assert.Equal("failed", status.State);
            Assert.Equal(message, status.ErrorMessage);
            Assert.True(status.WaitExpired);
            Assert.False(status.ExecutionPending);
            Assert.Equal(sessionId, Assert.Single(service.SessionManager.GetActiveSessions()).SessionId);
        }
        finally
        {
            GetPrivateField<ConcurrentDictionary<string, IExcelBatch>>(
                service.SessionManager, "_activeSessions").TryRemove(sessionId, out _);
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SerializePartialState_PreservesQueryFieldsAndExpiredOperation(bool hasQuery, bool waitExpired)
    {
        var query = hasQuery ? new PowerQueryPartialState
        {
            QueryName = "Retained query",
            QueryCreated = true,
            QueryExists = true,
            LoadMode = PowerQueryLoadMode.LoadToTable,
            TargetSheet = "Actual"
        } : null;
        var operation = new ExcelOperationStatus("running", 3.5, waitExpired, true, null, "shown")
        {
            Command = "powerquery.create"
        };
        var method = typeof(ExcelMcpService).GetMethod("SerializePartialState",
            BindingFlags.Static | BindingFlags.NonPublic);
        Assert.NotNull(method);
        var result = method.Invoke(null, [query, operation, null, null]);
        if (!hasQuery && !waitExpired)
        {
            Assert.Null(result);
            return;
        }
        using var json = JsonDocument.Parse(Assert.IsType<string>(result));
        Assert.Equal(hasQuery, json.RootElement.TryGetProperty("queryName", out var name));
        if (hasQuery)
        {
            Assert.Equal("Retained query", name.GetString());
            Assert.True(json.RootElement.GetProperty("queryCreated").GetBoolean());
            Assert.True(json.RootElement.GetProperty("queryExists").GetBoolean());
            Assert.Equal("load-to-table", json.RootElement.GetProperty("loadMode").GetString());
            Assert.Equal("Actual", json.RootElement.GetProperty("targetSheet").GetString());
        }
        Assert.Equal(waitExpired, json.RootElement.TryGetProperty("operation", out var status));
        if (waitExpired)
        {
            Assert.Equal("running", status.GetProperty("state").GetString());
            Assert.Equal(3.5, status.GetProperty("elapsedSeconds").GetDouble());
            Assert.True(status.GetProperty("waitExpired").GetBoolean());
            Assert.True(status.GetProperty("executionPending").GetBoolean());
            Assert.Equal("shown", status.GetProperty("windowReveal").GetString());
            Assert.Equal("powerquery.create", status.GetProperty("command").GetString());
        }
    }

    [Fact]
    public void CreateErrorResponse_PartialState_SurvivesContextAndSharedProtocolRoundTrip()
    {
        using var service = new ExcelMcpService();
        var state = new PowerQueryPartialState
        {
            QueryName = "Query",
            QueryCreated = true,
            QueryExists = true,
            LoadMode = PowerQueryLoadMode.LoadToTable,
            TargetSheet = "Actual",
            IsLoadedToDataModel = false
        };
        var error = new PowerQueryCommandException("Privacy failure with retained query", "Privacy",
            new InvalidOperationException("Original error"), state);
        var create = typeof(ExcelMcpService).GetMethod("CreateErrorResponse",
            BindingFlags.Instance | BindingFlags.NonPublic)!;
        var response = Assert.IsType<ServiceResponse>(create.Invoke(service, [error, null, null, null, null]));
        var attach = typeof(ExcelMcpService).GetMethod("AttachRequestContext",
            BindingFlags.Static | BindingFlags.NonPublic)!;
        response = Assert.IsType<ServiceResponse>(attach.Invoke(null,
            [new ServiceRequest { Command = "powerquery.create", SessionId = "session" }, response]));
        var shared = JsonSerializer.Deserialize<Sbroenne.ExcelMcp.ComInterop.ServiceClient.ServiceResponse>(
            JsonSerializer.Serialize(response, ServiceProtocol.JsonOptions), ServiceProtocol.JsonOptions)!;
        var roundTrip = ServiceResponse.FromShared(shared);

        Assert.False(roundTrip.Success);
        Assert.Equal("powerquery.create", roundTrip.Command);
        Assert.Equal("session", roundTrip.SessionId);
        Assert.Equal("Privacy", roundTrip.ErrorCategory);
        Assert.Equal("PowerQueryCommandException", roundTrip.ExceptionType);
        Assert.Equal(response.ErrorMessage, roundTrip.ErrorMessage);
        Assert.Equal(response.InnerError, roundTrip.InnerError);
        var json = JsonSerializer.Deserialize<JsonElement>(Assert.IsType<string>(roundTrip.PartialState));
        Assert.Equal("Query", json.GetProperty("queryName").GetString());
        Assert.True(json.GetProperty("queryCreated").GetBoolean());
        Assert.True(json.GetProperty("queryExists").GetBoolean());
        Assert.Equal("load-to-table", json.GetProperty("loadMode").GetString());
        Assert.Equal("Actual", json.GetProperty("targetSheet").GetString());
        Assert.False(json.GetProperty("isLoadedToDataModel").GetBoolean());
        Assert.False(json.TryGetProperty("mCodeUpdated", out _));
        Assert.False(json.TryGetProperty("stateInspectionError", out _));
    }

    [Theory]
    [InlineData("wrapped-com", "ComInterop")]
    [InlineData("cancelled", "Cancelled")]
    [InlineData("unknown", null)]
    [InlineData("prerequisite", "Prerequisite")]
    [InlineData("dependency", "DependencyUnavailable")]
    [InlineData("permissions", "Permissions")]
    public void CreateErrorResponse_PreservesKnownNestedCategory(string scenario, string? category)
    {
        using var service = new ExcelMcpService();
#pragma warning disable CA2201 // Synthetic exceptions exercise serialization, not COM behavior.
        Exception error = scenario switch
        {
            "wrapped-com" => new InvalidOperationException("Operation context",
                new TargetInvocationException(new COMException("Excel failure", unchecked((int)0x800A03EC)))),
            "cancelled" => new OperationCanceledException("Cancelled operation"),
            "prerequisite" => new OperationFailureException(OperationFailureCategory.Prerequisite, "Missing model"),
            "dependency" => new OperationFailureException(OperationFailureCategory.DependencyUnavailable, "Missing provider"),
            "permissions" => new OperationFailureException(OperationFailureCategory.Permissions, "Access blocked"),
            _ => new InvalidOperationException("Unknown condition")
        };
#pragma warning restore CA2201
        var method = typeof(ExcelMcpService).GetMethod("CreateErrorResponse",
            BindingFlags.Instance | BindingFlags.NonPublic)!;
        var response = Assert.IsType<ServiceResponse>(method.Invoke(service, [error, "vba.run", "session", null, null]));

        Assert.False(response.Success);
        Assert.Equal(category, response.ErrorCategory);
        Assert.Contains(error.Message, response.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal(error.GetType().Name, response.ExceptionType);
        Assert.Equal("vba.run", response.Command);
        Assert.Equal("session", response.SessionId);
        Assert.Null(response.PartialState);
        if (scenario == "wrapped-com")
        {
            Assert.Equal("0x800A03EC", response.HResult);
        }
    }

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
    /// Unknown command groups are rejected by the shared generated validator, the same
    /// way unknown actions and parameters are, and the error lists the valid groups.
    /// </summary>
    [Fact]
    public async Task ProcessAsync_UnknownCategory_ReturnsInvalidInputListingValidGroups()
    {
        // Arrange
        using var service = new ExcelMcpService();
        var request = new ServiceRequest { Command = "unknowncategory.someaction" };

        // Act
        var response = await service.ProcessAsync(request);

        // Assert
        Assert.False(response.Success);
        Assert.Equal("InvalidInput", response.ErrorCategory);
        Assert.NotNull(response.ErrorMessage);
        Assert.Contains("Unknown command group 'unknowncategory'", response.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains("Valid groups:", response.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains("worksheetstyle", response.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains("calculationmode", response.ErrorMessage, StringComparison.Ordinal);
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
        Assert.Equal("SessionNotFound", response.ErrorCategory);
    }

    [Fact]
    public async Task ProcessAsync_SessionCommandWithoutSessionId_ReturnsInvalidInputCategory()
    {
        using var service = new ExcelMcpService();

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "sheet.list"
        });

        Assert.False(response.Success);
        Assert.Equal("InvalidInput", response.ErrorCategory);
        Assert.Contains("sessionId", response.ErrorMessage, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ProcessAsync_SessionCreateForExistingFile_ReturnsConflictCategory()
    {
        using var service = new ExcelMcpService();
        var existingPath = Path.Combine(
            Path.GetTempPath(),
            $"{Guid.NewGuid():N}.xlsx");
        File.WriteAllText(existingPath, string.Empty);

        try
        {
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.create",
                Args = JsonSerializer.Serialize(new { filePath = existingPath })
            });

            Assert.False(response.Success);
            Assert.Equal("Conflict", response.ErrorCategory);
        }
        finally
        {
            File.Delete(existingPath);
        }
    }

    [Fact]
    public async Task ProcessAsync_SessionCloseValidatesArgumentsBeforeSessionId()
    {
        using var service = new ExcelMcpService();

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            Args = """{"save":"invalid"}"""
        });

        Assert.False(response.Success);
        Assert.Equal("InvalidInput", response.ErrorCategory);
        Assert.Contains("save", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("sessionId", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(
        "powerquery.refresh",
        """{"queryName":"Probe","TimeoutSeconds":null}""",
        "TimeoutSeconds")]
    [InlineData(
        "powerquery.refresh",
        """{"queryName":"Probe","unexpected":true}""",
        "unexpected")]
    [InlineData(
        "connection.refresh",
        """{"connectionName":"Probe","timeoutSeconds":0}""",
        "timeoutSeconds")]
    [InlineData(
        "powerquery.load-to",
        """{"queryName":"Probe","loadDestination":"worksheet","timeout":60}""",
        "timeout")]
    [InlineData(
        "powerquery.refresh",
        """{"queryName":null}""",
        "queryName")]
    [InlineData(
        "powerquery.evaluate",
        """{}""",
        "mCode")]
    [InlineData(
        "powerquery.evaluate",
        """{"mCode":"let Source = 1 in Source","mCodeFile":"missing.m"}""",
        "mCode")]
    [InlineData(
        "chartconfig.remove-series",
        """{"chartName":"Chart 1"}""",
        "seriesIndex")]
    [InlineData(
        "analysis.create-scenario",
        """{"sheetName":"Model","scenarioName":"Base","changingCells":"A1"}""",
        "values")]
    [InlineData(
        "table.toggle-totals",
        """{"tableName":"Probe"}""",
        "showTotals")]
    [InlineData(
        "powerquery.load-to",
        """{"queryName":"Probe","loadDestination":"worksheet","timeout":null}""",
        "timeout")]
    [InlineData(
        "window.set-position",
        """{"left":"not-a-number"}""",
        "left")]
    public async Task ProcessAsync_ValidatesGeneratedArgumentsBeforeSessionLookup(
        string command,
        string argsJson,
        string expectedParameter)
    {
        using var service = new ExcelMcpService();

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = command,
            SessionId = "missing-session",
            Args = argsJson
        });

        Assert.False(response.Success);
        Assert.Equal("InvalidInput", response.ErrorCategory);
        Assert.Contains(expectedParameter, response.ErrorMessage, StringComparison.Ordinal);
        Assert.DoesNotContain("session", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(
        """{"filePath":"C:\\missing-parent\\probe.xlsx","macroEnabled":"true"}""",
        "macroEnabled")]
    [InlineData(
        """{"filePath":"C:\\missing-parent\\probe.xlsx","macroEnabled":true}""",
        "macroEnabled")]
    public async Task ProcessAsync_ValidatesSessionCreateMacroEnabledBeforeFileCreation(
        string argsJson,
        string expectedParameter)
    {
        using var service = new ExcelMcpService();

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.create",
            Args = argsJson
        });

        Assert.False(response.Success);
        Assert.Equal("InvalidInput", response.ErrorCategory);
        Assert.Contains(expectedParameter, response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ServiceClient_DefaultRequestTimeout_ExceedsLargestPublicOperationTimeout()
    {
        Assert.True(
            ServiceClient.DefaultRequestTimeout > TimeSpan.FromSeconds(ParameterTransforms.MaximumTimeoutSeconds));
    }

    [Fact]
    public async Task ProcessAsync_ValidatesFileInputBeforeSessionLookup()
    {
        using var service = new ExcelMcpService();
        var missingPath = Path.Join(Path.GetTempPath(), $"{Guid.NewGuid():N}.m");

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "powerquery.evaluate",
            SessionId = "missing-session",
            Args = JsonSerializer.Serialize(new { mCodeFile = missingPath })
        });

        Assert.False(response.Success);
        Assert.Contains("not found", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(missingPath, response.ErrorMessage, StringComparison.Ordinal);
        Assert.DoesNotContain("session", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ProcessAsync_SessionCommandOnTimedOutSession_FailsFastBeforeExecutingBatch()
    {
        using var service = new ExcelMcpService();
        var batch = new FakeBatch { HasTimedOutOperation = true };
        const string sessionId = "timed-out-sheet-list";

        RegisterSession(service, sessionId, batch);

        try
        {
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "sheet.list",
                SessionId = sessionId
            });

            Assert.False(response.Success);
            Assert.Equal("SessionUnavailable", response.ErrorCategory);
            Assert.NotNull(response.ErrorMessage);
            Assert.Contains("expired", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("remains open", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, batch.ExecuteCalls);
        }
        finally
        {
            batch.HasTimedOutOperation = false;
        }
    }

    [Fact]
    public async Task ProcessAsync_ProcessWordInOrdinaryOperationError_PreservesHealthySession()
    {
        using var service = new ExcelMcpService();
        var batch = new FakeBatch
        {
            ExecuteException = new InvalidOperationException("Worksheet 'Processing' was not found.")
        };
        const string sessionId = "healthy-process-word";

        RegisterSession(service, sessionId, batch);

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "sheet.list",
            SessionId = sessionId
        });

        Assert.False(response.Success);
        Assert.NotEqual("ExcelProcessDied", response.ErrorCategory);
        Assert.Contains("Processing", response.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal(batch, GetPrivateField<SessionManager>(service, "_sessionManager").GetSession(sessionId));
        Assert.Equal(0, batch.DisposeCalls);
    }

    [Fact]
    public async Task ProcessAsync_QueuedTimeout_PreservesUsableSession()
    {
        using var service = new ExcelMcpService();
        var batch = new FakeBatch
        {
            ExecuteException = new TimeoutException("Expired before the Excel thread started the operation.")
        };
        const string sessionId = "queued-timeout";

        RegisterSession(service, sessionId, batch);

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "sheet.list",
            SessionId = sessionId
        });

        Assert.False(response.Success);
        Assert.Equal("Timeout", response.ErrorCategory);
        Assert.DoesNotContain("session has been closed", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(batch, GetPrivateField<SessionManager>(service, "_sessionManager").GetSession(sessionId));
        Assert.Equal(0, batch.DisposeCalls);
    }

    [Theory]
    [InlineData(unchecked((int)0x800706BA), "aggregate")]
    [InlineData(unchecked((int)0x800706BE), "aggregate")]
    [InlineData(unchecked((int)0x80010108), "aggregate")]
    [InlineData(unchecked((int)0x80010108), "nested")]
    [InlineData(unchecked((int)0x80010108), "cancelled")]
    [InlineData(unchecked((int)0x80010108), "timeout")]
    public async Task ProcessAsync_SecondaryAggregateDisconnect_PreservesLiveSession(int hresult, string wrapper)
    {
        using var service = new ExcelMcpService();
        var mutationError = new InvalidOperationException("Workbook mutation failed.");
#pragma warning disable CA2201 // Synthetic COM failures exercise Service transport without Excel.
        Exception cleanupError = new COMException("Cleanup disconnected from Excel.", hresult);
#pragma warning restore CA2201
        if (wrapper == "nested")
        {
            cleanupError = new TargetInvocationException(new AggregateException(
                new InvalidOperationException("Another cleanup failed."), cleanupError));
        }
        Exception error = new AggregateException(mutationError, cleanupError);
        error = wrapper switch
        {
            "nested" => new InvalidOperationException("Operation failed.", new TargetInvocationException(error)),
            "cancelled" => new OperationCanceledException("Operation cancelled.", error),
            "timeout" => new TimeoutException("Operation timed out.", error),
            _ => error
        };
        var batch = new FakeBatch { ExecuteException = error };
        const string sessionId = "secondary-aggregate-disconnect";
        RegisterSession(service, sessionId, batch);

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "sheet.list",
            SessionId = sessionId
        });

        Assert.False(response.Success);
        Assert.Equal("ExcelDisconnected", response.ErrorCategory);
        Assert.Equal($"0x{hresult:X8}", response.HResult);
        Assert.Equal(error.GetType().Name, response.ExceptionType);
        Assert.Equal("sheet.list", response.Command);
        Assert.Equal(sessionId, response.SessionId);
        Assert.Contains("still live", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(error.InnerException!.Message, response.InnerError);
        Assert.False(batch.HasTimedOutOperation);
        Assert.True(batch.IsAlive);
        Assert.Equal(1, batch.ExecuteCalls);
        Assert.Equal(0, batch.DisposeCalls);
        Assert.Equal(0, batch.SaveCalls);

        var listResponse = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
        Assert.True(listResponse.Success);
        Assert.NotNull(listResponse.Result);
        Assert.Contains(sessionId, listResponse.Result, StringComparison.Ordinal);

        var retryResponse = await service.ProcessAsync(new ServiceRequest
        {
            Command = "sheet.list",
            SessionId = sessionId
        });
        Assert.False(retryResponse.Success);
        Assert.Equal("ExcelDisconnected", retryResponse.ErrorCategory);
        Assert.Equal(2, batch.ExecuteCalls);
    }

    [Theory]
    [InlineData("aggregate")]
    [InlineData("cancelled")]
    [InlineData("timeout")]
    public async Task ProcessAsync_NonfatalAggregateOrCancellation_PreservesHealthySession(string wrapper)
    {
        using var service = new ExcelMcpService();
#pragma warning disable CA2201 // Synthetic COM failures exercise Service transport without Excel.
        Exception error = new AggregateException(
            new InvalidOperationException("Workbook mutation failed."),
            new COMException("Ordinary Excel error.", unchecked((int)0x800A03EC)));
#pragma warning restore CA2201
        error = wrapper switch
        {
            "cancelled" => new OperationCanceledException("Operation cancelled.", error),
            "timeout" => new TimeoutException("Operation timed out.", error),
            _ => error
        };
        var batch = new FakeBatch { ExecuteException = error };
        const string sessionId = "healthy-aggregate-error";
        RegisterSession(service, sessionId, batch);

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "sheet.list",
            SessionId = sessionId
        });

        Assert.False(response.Success);
        if (wrapper == "cancelled")
        {
            Assert.Equal("Cancelled", response.ErrorCategory);
        }
        else if (wrapper == "timeout")
        {
            Assert.Equal("Timeout", response.ErrorCategory);
        }
        Assert.NotEqual("ExcelProcessDied", response.ErrorCategory);
        Assert.Contains(error.Message, response.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal(error.GetType().Name, response.ExceptionType);
        Assert.Equal(batch, GetPrivateField<SessionManager>(service, "_sessionManager").GetSession(sessionId));
        Assert.Equal(1, batch.ExecuteCalls);
        Assert.Equal(0, batch.DisposeCalls);
        Assert.Equal(0, batch.SaveCalls);
    }

    [Fact]
    public void FatalDisconnectClassification_DoesNotInferDeathFromMessageText()
    {
        var method = typeof(ExcelMcpService).GetMethod(
            "TryGetFatalComHResult",
            BindingFlags.Static | BindingFlags.NonPublic)!;
        var error = new InvalidOperationException("RPC server is unavailable while validating input.");

        var result = method.Invoke(null, [error]);

        Assert.Null(result);
    }

    [Fact]
    public async Task ProcessAsync_SessionCloseSaveOnTimedOutSession_FailsFastBeforeSaving()
    {
        using var service = new ExcelMcpService();
        var batch = new FakeBatch { HasTimedOutOperation = true };
        const string sessionId = "timed-out-close-save";

        RegisterSession(service, sessionId, batch);

        try
        {
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.close",
                SessionId = sessionId,
                Args = JsonSerializer.Serialize(new { save = true }, ServiceProtocol.JsonOptions)
            });

            Assert.False(response.Success);
            Assert.NotNull(response.ErrorMessage);
            Assert.Contains("busy", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, batch.SaveCalls);
        }
        finally
        {
            batch.HasTimedOutOperation = false;
        }
    }

    [Fact]
    public async Task ProcessAsync_SessionCloseSaveOnHealthySession_SavesNormally()
    {
        using var service = new ExcelMcpService();
        var batch = new FakeBatch();
        const string sessionId = "healthy-close-save";

        RegisterSession(service, sessionId, batch);

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new { save = true }, ServiceProtocol.JsonOptions)
        });

        Assert.True(response.Success);
        Assert.Equal(1, batch.SaveCalls);
    }

    [Fact]
    public async Task ProcessAsync_SessionCloseSaveAfterRpcDisconnected_CleansSessionAndReturnsActionableError()
    {
        using var service = new ExcelMcpService();
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
        Assert.Contains("Dead-session tracking has been removed", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(1, batch.SaveCalls);
        Assert.Equal(1, batch.DisposeCalls);

        var listResponse = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
        Assert.True(listResponse.Success);
        Assert.NotNull(listResponse.Result);
        Assert.DoesNotContain(sessionId, listResponse.Result, StringComparison.Ordinal);
    }

    private static void RegisterSession(ExcelMcpService service, string sessionId, IExcelBatch batch)
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

    private sealed class FakeBatch : IExcelBatch, IExcelBatchRefreshState
    {
        public WorkbookRefreshState GetRefreshState() =>
            HasTimedOutOperation ? WorkbookRefreshState.Busy : WorkbookRefreshState.Ready;
        public string WorkbookPath { get; init; } = Path.Combine(Path.GetTempPath(), $"fake-batch-{Guid.NewGuid():N}.xlsx");
        public Microsoft.Extensions.Logging.ILogger Logger { get; } = NullLogger.Instance;
        public IReadOnlyDictionary<string, Excel.Workbook> Workbooks { get; } = new Dictionary<string, Excel.Workbook>();
        public bool HasTimedOutOperation { get; set; }
        public bool IsAlive { get; private set; } = true;
        public bool IsAliveAfterSaveException { get; init; } = true;
        public Exception? ExecuteException { get; init; }
        public Exception? SaveException { get; init; }
        public int ExecuteCalls { get; private set; }
        public int SaveCalls { get; private set; }
        public int DisposeCalls { get; private set; }
        public int? ExcelProcessId => 1234;
        public TimeSpan OperationTimeout => TimeSpan.FromSeconds(5);
        public bool IsExcelVisible => false;

        public Excel.Workbook GetWorkbook(string filePath) => throw new NotSupportedException();

        public void UpdateWorkbookPath(string workbookPath) => throw new NotSupportedException();

        public void Execute(Action<ExcelContext, CancellationToken> operation, CancellationToken cancellationToken = default)
        {
            ExecuteCalls++;
            if (ExecuteException != null)
            {
                throw ExecuteException;
            }

            throw new InvalidOperationException("Execute should not be called for this unavailable fake batch.");
        }

        public T Execute<T>(Func<ExcelContext, CancellationToken, T> operation, CancellationToken cancellationToken = default)
        {
            ExecuteCalls++;
            if (ExecuteException != null)
            {
                throw ExecuteException;
            }

            throw new InvalidOperationException("Execute should not be called for this unavailable fake batch.");
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
