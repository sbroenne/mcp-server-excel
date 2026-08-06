using System.Diagnostics;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Proves the safety handshake through the shared service dispatcher and a real Excel session.
/// </summary>
[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "Safety")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Medium")]
public sealed class SafetyWorkflowRealExcelTests : IClassFixture<TempDirectoryFixture>
{
    private readonly string _tempDirectory;

    public SafetyWorkflowRealExcelTests(TempDirectoryFixture fixture)
    {
        _tempDirectory = Path.Combine(fixture.TempDir, $"S-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDirectory);
    }

    [Fact]
    public async Task ReviewRequired_RangeMutation_ReviewsOnceAndExecutesExactlyOnce()
    {
        const string journalSecretSentinel = "JournalSecretCellValue";
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            "Safety",
            "range-review",
            _tempDirectory,
            ".xlsx");
        using var service = new ExcelMcpService(_tempDirectory);

        var create = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new { filePath = workbookPath, show = ShowExcel }, ServiceProtocol.JsonOptions)
        });
        Assert.True(create.Success, create.ErrorMessage);
        var sessionId = GetRequiredString(create.Result, "sessionId");

        var configure = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.configure-safety",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                reviewMode = "required",
                checkpointMode = "off",
                journalMode = "on",
                verificationMode = "on",
                abnormalShutdownPolicy = "discardWithRecoveryEvidence"
            }, ServiceProtocol.JsonOptions)
        });
        Assert.True(configure.Success, configure.ErrorMessage);

        var mutationArgs = JsonSerializer.Serialize(new
        {
            sheetName = "Sheet1",
            rangeAddress = "A1",
            values = new object?[][] { [journalSecretSentinel] }
        }, ServiceProtocol.JsonOptions);

        var blocked = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = mutationArgs
        });
        Assert.False(blocked.Success);
        Assert.Equal("ReviewRequired", blocked.ErrorCategory);

        var review = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = mutationArgs,
            ReviewOnly = true
        });
        Assert.True(review.Success, review.ErrorMessage);
        Assert.False(GetRequiredBoolean(review.Result, "executed"));
        var reviewId = GetRequiredString(review.Result, "reviewId");

        var before = await GetSingleCellAsync(service, sessionId);
        Assert.Equal(JsonValueKind.Null, before.ValueKind);

        var execute = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = mutationArgs,
            ReviewId = reviewId
        });
        Assert.True(execute.Success, execute.ErrorMessage);
        Assert.True(GetRequiredBoolean(execute.Result, "executed"));
        using (var executeReceipt = JsonDocument.Parse(execute.Result!))
        {
            var verification = executeReceipt.RootElement.GetProperty("verification");
            Assert.Equal("verified", verification.GetProperty("status").GetString());
            Assert.Equal(1, verification.GetProperty("changedCells").GetInt32());
        }

        var after = await GetSingleCellAsync(service, sessionId);
        Assert.Equal(journalSecretSentinel, after.GetString());
        PauseIfVisible();

        var reuse = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = mutationArgs,
            ReviewId = reviewId
        });
        Assert.False(reuse.Success);
        Assert.Equal("ReviewConsumed", reuse.ErrorCategory);

        var afterReuse = await GetSingleCellAsync(service, sessionId);
        Assert.Equal(journalSecretSentinel, afterReuse.GetString());

        var journal = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.journal",
            SessionId = sessionId
        });
        Assert.True(journal.Success, journal.ErrorMessage);
        Assert.DoesNotContain(journalSecretSentinel, journal.Result, StringComparison.Ordinal);
        using (var journalDocument = JsonDocument.Parse(journal.Result!))
        {
            var operation = Assert.Single(journalDocument.RootElement.GetProperty("operations").EnumerateArray());
            var argumentSummary = operation.GetProperty("argumentSummary");
            Assert.Equal(3, argumentSummary.GetProperty("parameterCount").GetInt32());
            Assert.Equal(2, argumentSummary.GetProperty("stringCount").GetInt32());
            Assert.Equal(1, argumentSummary.GetProperty("arrayCount").GetInt32());
        }

        _ = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = sessionId,
            Args = "{\"save\":false}"
        });
    }

    [Fact]
    public async Task ReviewRequired_RangeCopy_ReportsExpandedTargetAndRejectsChangedSource()
    {
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            "Safety",
            "copy-review",
            _tempDirectory,
            ".xlsx");
        using var service = new ExcelMcpService(_tempDirectory);
        var sessionId = await CreateSessionAsync(service, workbookPath);

        var seed = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = "{\"sheetName\":\"Sheet1\",\"rangeAddress\":\"A1:B1\",\"values\":[[10,20]]}"
        });
        Assert.True(seed.Success, seed.ErrorMessage);
        await ConfigureRequiredReviewAsync(service, sessionId);

        const string copyArgs = "{\"sourceSheet\":\"Sheet1\",\"sourceRange\":\"A1:B1\",\"targetSheet\":\"Sheet1\",\"targetRange\":\"D2\"}";
        var review = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.copy-values",
            SessionId = sessionId,
            Args = copyArgs,
            ReviewOnly = true
        });
        Assert.True(review.Success, review.ErrorMessage);
        using (var reviewDocument = JsonDocument.Parse(review.Result!))
        {
            var affected = reviewDocument.RootElement.GetProperty("affected");
            Assert.Equal(["Sheet1"], affected.GetProperty("sheets").EnumerateArray().Select(item => item.GetString()));
            Assert.Equal(["Sheet1!$D$2:$E$2"], affected.GetProperty("ranges").EnumerateArray().Select(item => item.GetString()));
            Assert.Empty(affected.GetProperty("objects").EnumerateArray());
        }

        var reviewId = GetRequiredString(review.Result, "reviewId");
        var batch = service.SessionManager.GetSession(sessionId)
            ?? throw new InvalidOperationException("Expected the reviewed Excel session to remain available.");
        SetCellValue(batch, "Sheet1", "A1", 99);

        var stale = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.copy-values",
            SessionId = sessionId,
            Args = copyArgs,
            ReviewId = reviewId
        });
        Assert.False(stale.Success);
        Assert.Equal("ReviewStale", stale.ErrorCategory);

        var freshReview = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.copy-values",
            SessionId = sessionId,
            Args = copyArgs,
            ReviewOnly = true
        });
        Assert.True(freshReview.Success, freshReview.ErrorMessage);
        var execute = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.copy-values",
            SessionId = sessionId,
            Args = copyArgs,
            ReviewId = GetRequiredString(freshReview.Result, "reviewId")
        });
        Assert.True(execute.Success, execute.ErrorMessage);
        using (var receipt = JsonDocument.Parse(execute.Result!))
        {
            Assert.Equal("verified", receipt.RootElement.GetProperty("verification").GetProperty("status").GetString());
        }

        var copied = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.get-values",
            SessionId = sessionId,
            Args = "{\"sheetName\":\"Sheet1\",\"rangeAddress\":\"D2:E2\"}"
        });
        Assert.True(copied.Success, copied.ErrorMessage);
        using (var copiedDocument = JsonDocument.Parse(copied.Result!))
        {
            var values = copiedDocument.RootElement.GetProperty("values")[0];
            Assert.Equal(99, values[0].GetInt32());
            Assert.Equal(20, values[1].GetInt32());
        }

        PauseIfVisible();
        _ = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = sessionId,
            Args = "{\"save\":false}"
        });
    }

    [Fact]
    public async Task ReviewRequired_LargeWorkbookRangeMutation_VerifiesOnlyTheTargetScope()
    {
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            "Safety",
            "target-scope",
            _tempDirectory,
            ".xlsx");
        using var service = new ExcelMcpService(_tempDirectory);
        var sessionId = await CreateSessionAsync(service, workbookPath);

        var rows = Enumerable.Range(1, 2_501)
            .Select(row => Enumerable.Range(1, 10)
                .Select(column => (object?)(row * 100 + column))
                .ToList())
            .ToList();
        var seed = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                sheetName = "Sheet1",
                rangeAddress = "A1:J2501",
                values = rows
            }, ServiceProtocol.JsonOptions)
        });
        Assert.True(seed.Success, seed.ErrorMessage);

        await ConfigureRequiredReviewAsync(service, sessionId);
        const string mutationArgs = "{\"sheetName\":\"Sheet1\",\"rangeAddress\":\"B2\",\"values\":[[999999]]}";
        var review = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = mutationArgs,
            ReviewOnly = true
        });
        Assert.True(review.Success, review.ErrorMessage);
        var reviewId = GetRequiredString(review.Result, "reviewId");

        var execute = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = mutationArgs,
            ReviewId = reviewId
        });
        Assert.True(execute.Success, execute.ErrorMessage);
        using (var receipt = JsonDocument.Parse(execute.Result!))
        {
            var verification = receipt.RootElement.GetProperty("verification");
            Assert.Equal("verified", verification.GetProperty("status").GetString());
            Assert.Equal(1, verification.GetProperty("changedCells").GetInt32());
        }

        var close = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = sessionId,
            Args = "{\"save\":false}"
        });
        Assert.True(close.Success, close.ErrorMessage);
    }

    [Fact]
    public async Task UnavailablePreExecutionJournal_BlocksMutationAndKeepsHealthySessionUsable()
    {
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            "Safety",
            "journal-down",
            _tempDirectory,
            ".xlsx");
        using var service = new ExcelMcpService(_tempDirectory);
        var sessionId = await CreateSessionAsync(service, workbookPath);

        var configure = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.configure-safety",
            SessionId = sessionId,
            Args = "{\"reviewMode\":\"off\",\"checkpointMode\":\"off\",\"journalMode\":\"on\",\"verificationMode\":\"off\",\"abnormalShutdownPolicy\":\"discardWithRecoveryEvidence\"}"
        });
        Assert.True(configure.Success, configure.ErrorMessage);

        PauseIfVisible();
        var journalDirectory = Path.Combine(_tempDirectory, "journal");
        Directory.Delete(journalDirectory, recursive: true);
        File.WriteAllText(journalDirectory, "synthetic journal outage");

        var mutation = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = "{\"sheetName\":\"Sheet1\",\"rangeAddress\":\"A1\",\"values\":[[99]]}"
        });
        Assert.False(mutation.Success);
        Assert.NotNull(mutation.ExceptionType);

        var after = await GetSingleCellAsync(service, sessionId);
        Assert.Equal(JsonValueKind.Null, after.ValueKind);
        PauseIfVisible();

        var close = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = sessionId,
            Args = "{\"save\":false}"
        });
        Assert.True(close.Success, close.ErrorMessage);
    }

    [Fact]
    public async Task RequiredCheckpoint_RangeMutation_JournalsAndRecoversPreWriteState()
    {
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            "Safety",
            "checkpoint",
            _tempDirectory,
            ".xlsx");
        using var service = new ExcelMcpService(_tempDirectory);
        var sessionId = await CreateSessionAsync(service, workbookPath);

        var configure = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.configure-safety",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                reviewMode = "required",
                checkpointMode = "required",
                journalMode = "on",
                verificationMode = "on",
                abnormalShutdownPolicy = "discardWithRecoveryEvidence"
            }, ServiceProtocol.JsonOptions)
        });
        Assert.True(configure.Success, configure.ErrorMessage);

        const string mutationArgs = "{\"sheetName\":\"Sheet1\",\"rangeAddress\":\"A1\",\"values\":[[99]]}";
        var review = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = mutationArgs,
            ReviewOnly = true
        });
        Assert.True(review.Success, review.ErrorMessage);
        var reviewId = GetRequiredString(review.Result, "reviewId");

        var execute = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = mutationArgs,
            ReviewId = reviewId
        });
        Assert.True(execute.Success, execute.ErrorMessage);

        using var receipt = JsonDocument.Parse(execute.Result!);
        var checkpoint = receipt.RootElement.GetProperty("checkpoint");
        Assert.True(checkpoint.GetProperty("created").GetBoolean());
        var checkpointPath = checkpoint.GetProperty("path").GetString();
        Assert.NotNull(checkpointPath);
        Assert.True(File.Exists(checkpointPath));
        Assert.True(new FileInfo(checkpointPath).Length > 0);
        Assert.Equal(64, checkpoint.GetProperty("sha256").GetString()?.Length);
        Assert.Equal("verified", receipt.RootElement.GetProperty("verification").GetProperty("status").GetString());
        Assert.Equal(1, receipt.RootElement.GetProperty("verification").GetProperty("changedCells").GetInt32());
        PauseIfVisible();

        var journal = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.journal",
            SessionId = sessionId
        });
        Assert.True(journal.Success, journal.ErrorMessage);
        using var journalDocument = JsonDocument.Parse(journal.Result!);
        var transitions = journalDocument.RootElement.GetProperty("operations")[0].GetProperty("transitions")
            .EnumerateArray()
            .Select(item => item.GetProperty("state").GetString() ?? string.Empty)
            .ToArray();
        Assert.Equal<string>(["reviewed", "checkpointReserved", "checkpointCreated", "started", "completed", "verified"], transitions);

        var recoveries = await service.ProcessAsync(new ServiceRequest { Command = "recovery.list" });
        Assert.True(recoveries.Success, recoveries.ErrorMessage);
        using var recoveriesDocument = JsonDocument.Parse(recoveries.Result!);
        var recovery = recoveriesDocument.RootElement.GetProperty("recoveries")[0];
        var recoveryId = recovery.GetProperty("recoveryId").GetString();
        Assert.NotNull(recoveryId);
        Assert.True(recovery.GetProperty("available").GetBoolean());

        var recover = await service.ProcessAsync(new ServiceRequest
        {
            Command = "recovery.recover",
            Args = JsonSerializer.Serialize(new { recoveryId, show = ShowExcel }, ServiceProtocol.JsonOptions)
        });
        Assert.True(recover.Success, recover.ErrorMessage);
        var recoveredSessionId = GetRequiredString(recover.Result, "sessionId");

        var recoveredValue = await GetSingleCellAsync(service, recoveredSessionId);
        Assert.Equal(JsonValueKind.Null, recoveredValue.ValueKind);
        PauseIfVisible();

        _ = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = recoveredSessionId,
            Args = "{\"save\":false}"
        });
        _ = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = sessionId,
            Args = "{\"save\":false}"
        });
    }

    [Fact]
    public async Task AbnormalShutdown_DiscardPolicy_DoesNotAutoSavePartialWorkbookAndKeepsRecovery()
    {
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            "Safety",
            "shutdown",
            _tempDirectory,
            ".xlsx");
        var service = new ExcelMcpService(_tempDirectory);
        var sessionId = await CreateSessionAsync(service, workbookPath);

        var configure = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.configure-safety",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                reviewMode = "off",
                checkpointMode = "required",
                journalMode = "on",
                verificationMode = "on",
                abnormalShutdownPolicy = "discardWithRecoveryEvidence"
            }, ServiceProtocol.JsonOptions)
        });
        Assert.True(configure.Success, configure.ErrorMessage);

        var mutation = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = "{\"sheetName\":\"Sheet1\",\"rangeAddress\":\"A1\",\"values\":[[7]]}"
        });
        Assert.True(mutation.Success, mutation.ErrorMessage);
        PauseIfVisible();

        service.Dispose();

        using var restartedService = new ExcelMcpService(_tempDirectory);
        var reopened = await restartedService.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new { filePath = workbookPath, show = ShowExcel }, ServiceProtocol.JsonOptions)
        });
        Assert.True(reopened.Success, reopened.ErrorMessage);
        var reopenedSessionId = GetRequiredString(reopened.Result, "sessionId");
        var originalValue = await GetSingleCellAsync(restartedService, reopenedSessionId);
        Assert.Equal(JsonValueKind.Null, originalValue.ValueKind);
        PauseIfVisible();

        var recoveries = await restartedService.ProcessAsync(new ServiceRequest { Command = "recovery.list" });
        Assert.True(recoveries.Success, recoveries.ErrorMessage);
        using var recoveryDocument = JsonDocument.Parse(recoveries.Result!);
        Assert.True(recoveryDocument.RootElement.GetProperty("count").GetInt32() >= 1);

        _ = await restartedService.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = reopenedSessionId,
            Args = "{\"save\":false}"
        });
    }

    [Fact]
    public async Task ExcelProcessDeath_PreflightRecordsDurableEvidenceAndKeepsCheckpointRecoverable()
    {
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            "Safety",
            "death-preflight",
            _tempDirectory,
            ".xlsx");
        using var service = new ExcelMcpService(_tempDirectory);
        var sessionId = await CreateSessionAsync(service, workbookPath);

        var configure = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.configure-safety",
            SessionId = sessionId,
            Args = "{\"reviewMode\":\"off\",\"checkpointMode\":\"required\",\"journalMode\":\"on\",\"verificationMode\":\"on\",\"abnormalShutdownPolicy\":\"discardWithRecoveryEvidence\"}"
        });
        Assert.True(configure.Success, configure.ErrorMessage);

        var mutation = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = "{\"sheetName\":\"Sheet1\",\"rangeAddress\":\"A1\",\"values\":[[123]]}"
        });
        Assert.True(mutation.Success, mutation.ErrorMessage);
        PauseIfVisible();

        var batch = service.SessionManager.GetSession(sessionId);
        Assert.NotNull(batch);
        Assert.NotNull(batch.ExcelProcessId);
        using (var excelProcess = Process.GetProcessById(batch.ExcelProcessId.Value))
        {
            excelProcess.Kill(entireProcessTree: true);
            Assert.True(excelProcess.WaitForExit(10_000));
        }

        var afterDeath = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.preflight",
            SessionId = sessionId
        });
        Assert.False(afterDeath.Success);
        Assert.Equal("ExcelProcessDied", afterDeath.ErrorCategory);

        var sessionsAfterDeath = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
        Assert.True(sessionsAfterDeath.Success, sessionsAfterDeath.ErrorMessage);
        Assert.DoesNotContain(sessionId, sessionsAfterDeath.Result ?? string.Empty, StringComparison.Ordinal);

        var journal = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.journal",
            SessionId = sessionId
        });
        Assert.True(journal.Success, journal.ErrorMessage);
        using var journalDocument = JsonDocument.Parse(journal.Result!);
        var states = journalDocument.RootElement.GetProperty("operations")[0].GetProperty("transitions")
            .EnumerateArray()
            .Select(item => item.GetProperty("state").GetString() ?? string.Empty)
            .ToArray();
        Assert.Contains("excelProcessDied", states, StringComparer.Ordinal);

        var recoveries = await service.ProcessAsync(new ServiceRequest { Command = "recovery.list" });
        Assert.True(recoveries.Success, recoveries.ErrorMessage);
        using var recoveryDocument = JsonDocument.Parse(recoveries.Result!);
        Assert.True(recoveryDocument.RootElement.GetProperty("recoveries")[0].GetProperty("available").GetBoolean());
    }

    [Fact]
    public async Task ExcelProcessDeath_SessionListRecordsEvidenceBeforeAutomaticCleanup()
    {
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            "Safety",
            "death-list",
            _tempDirectory,
            ".xlsx");
        using var service = new ExcelMcpService(_tempDirectory);
        var sessionId = await CreateSessionAsync(service, workbookPath);

        var configure = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.configure-safety",
            SessionId = sessionId,
            Args = "{\"reviewMode\":\"off\",\"checkpointMode\":\"required\",\"journalMode\":\"on\",\"verificationMode\":\"on\",\"abnormalShutdownPolicy\":\"discardWithRecoveryEvidence\"}"
        });
        Assert.True(configure.Success, configure.ErrorMessage);

        var mutation = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = "{\"sheetName\":\"Sheet1\",\"rangeAddress\":\"A1\",\"values\":[[456]]}"
        });
        Assert.True(mutation.Success, mutation.ErrorMessage);

        var batch = service.SessionManager.GetSession(sessionId);
        Assert.NotNull(batch);
        Assert.NotNull(batch.ExcelProcessId);
        using (var excelProcess = Process.GetProcessById(batch.ExcelProcessId.Value))
        {
            excelProcess.Kill(entireProcessTree: true);
            Assert.True(excelProcess.WaitForExit(10_000));
        }

        var list = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
        Assert.True(list.Success, list.ErrorMessage);
        Assert.DoesNotContain(sessionId, list.Result ?? string.Empty, StringComparison.Ordinal);

        var journal = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.journal",
            SessionId = sessionId
        });
        Assert.True(journal.Success, journal.ErrorMessage);
        using var journalDocument = JsonDocument.Parse(journal.Result!);
        var states = journalDocument.RootElement.GetProperty("operations")[0].GetProperty("transitions")
            .EnumerateArray()
            .Select(item => item.GetProperty("state").GetString() ?? string.Empty)
            .ToArray();
        Assert.Contains("excelProcessDied", states, StringComparer.Ordinal);
    }

    [Fact]
    public async Task ExcelProcessDeath_ExplicitCloseRecordsEvidenceBeforeCleanup()
    {
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            "Safety",
            "death-close",
            _tempDirectory,
            ".xlsx");
        using var service = new ExcelMcpService(_tempDirectory);
        var sessionId = await CreateSessionAsync(service, workbookPath);

        var configure = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.configure-safety",
            SessionId = sessionId,
            Args = "{\"reviewMode\":\"off\",\"checkpointMode\":\"required\",\"journalMode\":\"on\",\"verificationMode\":\"on\",\"abnormalShutdownPolicy\":\"discardWithRecoveryEvidence\"}"
        });
        Assert.True(configure.Success, configure.ErrorMessage);

        var mutation = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = "{\"sheetName\":\"Sheet1\",\"rangeAddress\":\"A1\",\"values\":[[789]]}"
        });
        Assert.True(mutation.Success, mutation.ErrorMessage);

        var batch = service.SessionManager.GetSession(sessionId);
        Assert.NotNull(batch);
        Assert.NotNull(batch.ExcelProcessId);
        using (var excelProcess = Process.GetProcessById(batch.ExcelProcessId.Value))
        {
            excelProcess.Kill(entireProcessTree: true);
            Assert.True(excelProcess.WaitForExit(10_000));
        }

        var close = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = sessionId,
            Args = "{\"save\":false}"
        });
        Assert.False(close.Success);
        Assert.Equal("ExcelProcessDied", close.ErrorCategory);

        var journal = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.journal",
            SessionId = sessionId
        });
        Assert.True(journal.Success, journal.ErrorMessage);
        using var journalDocument = JsonDocument.Parse(journal.Result!);
        var states = journalDocument.RootElement.GetProperty("operations")[0].GetProperty("transitions")
            .EnumerateArray()
            .Select(item => item.GetProperty("state").GetString() ?? string.Empty)
            .ToArray();
        Assert.Contains("excelProcessDied", states, StringComparer.Ordinal);
    }

    [Fact]
    public async Task StructuralReview_BecomesStaleWhenSheetTableNameOrChartChanges()
    {
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            "Safety",
            "structure-stale",
            _tempDirectory,
            ".xlsx");
        using var service = new ExcelMcpService(_tempDirectory);
        var sessionId = await CreateSessionAsync(service, workbookPath);

        await SeedStructuralObjectsAsync(service, sessionId);
        await ConfigureRequiredReviewAsync(service, sessionId);

        await ReviewThenAssertStaleAsync(
            service, sessionId, "table.rename", "{\"tableName\":\"Sales\",\"newName\":\"SalesRenamed\"}", "table:Sales",
            batch => RenameTable(batch, "Sales", "SalesIntervening"));
        await ReviewThenAssertStaleAsync(
            service, sessionId, "namedrange.update", "{\"name\":\"TaxRate\",\"reference\":\"Sheet1!$A$2\"}", "name:TaxRate",
            batch => UpdateNameReference(batch, "TaxRate", "=Sheet1!$A$3"));
        await ReviewThenAssertStaleAsync(
            service, sessionId, "chart.move", "{\"chartName\":\"SalesChart\",\"left\":100}", "chart:SalesChart",
            batch => MoveChart(batch, "SalesChart", 25));
        await ReviewThenAssertStaleAsync(
            service, sessionId, "sheet.create", "{\"sheetName\":\"ReviewedSheet\"}", "worksheet:ReviewedSheet",
            batch => AddWorksheet(batch, "InterveningSheet"));

        PauseIfVisible();
        _ = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = sessionId,
            Args = "{\"save\":false}"
        });
    }

    [Fact]
    public async Task ExactRangeReview_BecomesStaleWhenUnrelatedSameCountStructureChanges()
    {
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            "Safety",
            "range-stale",
            _tempDirectory,
            ".xlsx");
        using var service = new ExcelMcpService(_tempDirectory);
        var sessionId = await CreateSessionAsync(service, workbookPath);

        await SeedStructuralObjectsAsync(service, sessionId);
        await ConfigureRequiredReviewAsync(service, sessionId);

        await ReviewExactRangeThenAssertStaleAsync(
            service,
            sessionId,
            batch => RenameTable(batch, "Sales", "SalesIntervening"));
        await ReviewExactRangeThenAssertStaleAsync(
            service,
            sessionId,
            batch => UpdateNameReference(batch, "TaxRate", "=Sheet1!$A$3"));
        await ReviewExactRangeThenAssertStaleAsync(
            service,
            sessionId,
            batch => MoveChart(batch, "SalesChart", 25));

        PauseIfVisible();
        _ = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = sessionId,
            Args = "{\"save\":false}"
        });
    }

    [Fact]
    public async Task RangeReview_EmptySheetName_ResolvesWorkbookDefinedNameAndVerifiesExactCells()
    {
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            "Safety",
            "named-range",
            _tempDirectory,
            ".xlsx");
        using var service = new ExcelMcpService(_tempDirectory);
        var sessionId = await CreateSessionAsync(service, workbookPath);

        var createName = await service.ProcessAsync(new ServiceRequest
        {
            Command = "namedrange.create",
            SessionId = sessionId,
            Args = "{\"name\":\"InputCell\",\"reference\":\"Sheet1!$A$1\"}"
        });
        Assert.True(createName.Success, createName.ErrorMessage);
        await ConfigureRequiredReviewAsync(service, sessionId);

        const string args = "{\"sheetName\":\"\",\"rangeAddress\":\"InputCell\",\"values\":[[42]]}";
        var review = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = args,
            ReviewOnly = true
        });
        Assert.True(review.Success, review.ErrorMessage);
        using var reviewDocument = JsonDocument.Parse(review.Result!);
        Assert.Contains("Sheet1", reviewDocument.RootElement.GetProperty("affected").GetProperty("sheets")
            .EnumerateArray().Select(item => item.GetString()));
        var reviewId = reviewDocument.RootElement.GetProperty("reviewId").GetString();
        Assert.NotNull(reviewId);

        var execute = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = args,
            ReviewId = reviewId
        });
        Assert.True(execute.Success, execute.ErrorMessage);
        using var executeDocument = JsonDocument.Parse(execute.Result!);
        var verification = executeDocument.RootElement.GetProperty("verification");
        Assert.Equal("verified", verification.GetProperty("status").GetString());
        Assert.Contains("Sheet1", verification.GetProperty("scope").GetProperty("sheets")
            .EnumerateArray().Select(item => item.GetString()));
        Assert.Contains("Sheet1!$A$1", verification.GetProperty("scope").GetProperty("ranges")
            .EnumerateArray().Select(item => item.GetString()));

        PauseIfVisible();
        _ = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = sessionId,
            Args = "{\"save\":false}"
        });
    }

    [Fact]
    public async Task StructuralReview_ExecutionIsPartiallyVerifiedForSheetTableNameAndChart()
    {
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            "Safety",
            "structure-verify",
            _tempDirectory,
            ".xlsx");
        using var service = new ExcelMcpService(_tempDirectory);
        var sessionId = await CreateSessionAsync(service, workbookPath);

        await SeedStructuralObjectsAsync(service, sessionId);
        await ConfigureRequiredReviewAsync(service, sessionId);

        await ReviewThenAssertPartiallyVerifiedAsync(
            service, sessionId, "table.rename", "{\"tableName\":\"Sales\",\"newName\":\"SalesRenamed\"}");
        await ReviewThenAssertPartiallyVerifiedAsync(
            service, sessionId, "namedrange.update", "{\"name\":\"TaxRate\",\"reference\":\"Sheet1!$A$2\"}");
        await ReviewThenAssertPartiallyVerifiedAsync(
            service, sessionId, "chart.move", "{\"chartName\":\"SalesChart\",\"left\":100}");
        await ReviewThenAssertPartiallyVerifiedAsync(
            service, sessionId, "sheet.create", "{\"sheetName\":\"ReviewedSheet\"}");

        PauseIfVisible();
        _ = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = sessionId,
            Args = "{\"save\":false}"
        });
    }

    private static async Task SeedStructuralObjectsAsync(ExcelMcpService service, string sessionId)
    {
        var values = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = "{\"sheetName\":\"Sheet1\",\"rangeAddress\":\"A1:B3\",\"values\":[[\"Category\",\"Value\"],[\"A\",1],[\"B\",2]]}"
        });
        Assert.True(values.Success, values.ErrorMessage);

        var table = await service.ProcessAsync(new ServiceRequest
        {
            Command = "table.create",
            SessionId = sessionId,
            Args = "{\"sheetName\":\"Sheet1\",\"tableName\":\"Sales\",\"rangeAddress\":\"A1:B3\"}"
        });
        Assert.True(table.Success, table.ErrorMessage);

        var name = await service.ProcessAsync(new ServiceRequest
        {
            Command = "namedrange.create",
            SessionId = sessionId,
            Args = "{\"name\":\"TaxRate\",\"reference\":\"Sheet1!$B$2\"}"
        });
        Assert.True(name.Success, name.ErrorMessage);

        var batch = service.SessionManager.GetSession(sessionId)
            ?? throw new InvalidOperationException("Expected the seeded Excel session to remain available.");
        AddChart(batch, "SalesChart");
    }

    private static async Task ConfigureRequiredReviewAsync(ExcelMcpService service, string sessionId)
    {
        var configure = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.configure-safety",
            SessionId = sessionId,
            Args = "{\"reviewMode\":\"required\",\"checkpointMode\":\"off\",\"journalMode\":\"off\",\"verificationMode\":\"on\"}"
        });
        Assert.True(configure.Success, configure.ErrorMessage);
    }

    private static async Task ReviewThenAssertStaleAsync(
        ExcelMcpService service, string sessionId, string command, string args, string expectedObject, Action<IExcelBatch> mutate)
    {
        var review = await service.ProcessAsync(new ServiceRequest
        {
            Command = command,
            SessionId = sessionId,
            Args = args,
            ReviewOnly = true
        });
        Assert.True(review.Success, review.ErrorMessage);
        using var reviewDocument = JsonDocument.Parse(review.Result!);
        Assert.Contains(expectedObject, reviewDocument.RootElement.GetProperty("affected").GetProperty("objects")
            .EnumerateArray().Select(item => item.GetString()));
        var reviewId = reviewDocument.RootElement.GetProperty("reviewId").GetString();
        Assert.NotNull(reviewId);

        var batch = service.SessionManager.GetSession(sessionId)
            ?? throw new InvalidOperationException("Expected the reviewed Excel session to remain available.");
        mutate(batch);
        var execute = await service.ProcessAsync(new ServiceRequest
        {
            Command = command,
            SessionId = sessionId,
            Args = args,
            ReviewId = reviewId
        });
        Assert.False(execute.Success);
        Assert.Equal("ReviewStale", execute.ErrorCategory);
    }

    private static async Task ReviewExactRangeThenAssertStaleAsync(
        ExcelMcpService service,
        string sessionId,
        Action<IExcelBatch> mutate)
    {
        const string command = "range.set-values";
        const string args = "{\"sheetName\":\"Sheet1\",\"rangeAddress\":\"B3\",\"values\":[[99]]}";
        var review = await service.ProcessAsync(new ServiceRequest
        {
            Command = command,
            SessionId = sessionId,
            Args = args,
            ReviewOnly = true
        });
        Assert.True(review.Success, review.ErrorMessage);
        using var reviewDocument = JsonDocument.Parse(review.Result!);
        var affected = reviewDocument.RootElement.GetProperty("affected");
        Assert.Empty(affected.GetProperty("objects").EnumerateArray());
        Assert.Contains(
            "Sheet1!$B$3",
            affected.GetProperty("ranges").EnumerateArray().Select(item => item.GetString()));
        var reviewId = reviewDocument.RootElement.GetProperty("reviewId").GetString();
        Assert.NotNull(reviewId);

        var batch = service.SessionManager.GetSession(sessionId)
            ?? throw new InvalidOperationException("Expected the reviewed Excel session to remain available.");
        mutate(batch);

        var execute = await service.ProcessAsync(new ServiceRequest
        {
            Command = command,
            SessionId = sessionId,
            Args = args,
            ReviewId = reviewId
        });
        Assert.False(execute.Success);
        Assert.Equal("ReviewStale", execute.ErrorCategory);
    }

    private static async Task ReviewThenAssertPartiallyVerifiedAsync(
        ExcelMcpService service, string sessionId, string command, string args)
    {
        var review = await service.ProcessAsync(new ServiceRequest
        {
            Command = command,
            SessionId = sessionId,
            Args = args,
            ReviewOnly = true
        });
        Assert.True(review.Success, review.ErrorMessage);
        var reviewId = GetRequiredString(review.Result, "reviewId");

        var execute = await service.ProcessAsync(new ServiceRequest
        {
            Command = command,
            SessionId = sessionId,
            Args = args,
            ReviewId = reviewId
        });
        Assert.True(execute.Success, execute.ErrorMessage);
        using var receipt = JsonDocument.Parse(execute.Result!);
        Assert.Equal("partiallyVerified", receipt.RootElement.GetProperty("verification").GetProperty("status").GetString());
    }

    private static void AddChart(IExcelBatch batch, string name) => batch.Execute((context, _) =>
    {
        dynamic? worksheet = null;
        dynamic? shapes = null;
        dynamic? shape = null;
        try
        {
            worksheet = context.Book.Worksheets.Item["Sheet1"];
            shapes = worksheet.Shapes;
            shape = shapes.AddChart(51, 0, 0, 200, 120);
            shape.Name = name;
        }
        finally
        {
            ComUtilities.Release(ref shape);
            ComUtilities.Release(ref shapes);
            ComUtilities.Release(ref worksheet);
        }
    });

    private static void RenameTable(IExcelBatch batch, string oldName, string newName) => batch.Execute((context, _) =>
    {
        dynamic? table = null;
        try
        {
            table = context.Book.Worksheets.Item["Sheet1"].ListObjects.Item[oldName];
            table.Name = newName;
        }
        finally { ComUtilities.Release(ref table); }
    });

    private static void UpdateNameReference(IExcelBatch batch, string name, string reference) => batch.Execute((context, _) =>
    {
        dynamic? definedName = null;
        try
        {
            definedName = context.Book.Names.Item(name);
            definedName.RefersTo = reference;
        }
        finally { ComUtilities.Release(ref definedName); }
    });

    private static void MoveChart(IExcelBatch batch, string name, double left) => batch.Execute((context, _) =>
    {
        dynamic? worksheet = null;
        dynamic? charts = null;
        dynamic? chart = null;
        try
        {
            worksheet = context.Book.Worksheets.Item["Sheet1"];
            charts = worksheet.ChartObjects();
            chart = charts.Item(name);
            chart.Left = left;
        }
        finally
        {
            ComUtilities.Release(ref chart);
            ComUtilities.Release(ref charts);
            ComUtilities.Release(ref worksheet);
        }
    });

    private static void AddWorksheet(IExcelBatch batch, string name) => batch.Execute((context, _) =>
    {
        dynamic? worksheets = null;
        dynamic? worksheet = null;
        try
        {
            worksheets = context.Book.Worksheets;
            worksheet = worksheets.Add();
            worksheet.Name = name;
        }
        finally
        {
            ComUtilities.Release(ref worksheet);
            ComUtilities.Release(ref worksheets);
        }
    });

    private static void SetCellValue(IExcelBatch batch, string sheetName, string rangeAddress, object value) =>
        batch.Execute((context, _) =>
        {
            dynamic? worksheet = null;
            dynamic? range = null;
            try
            {
                worksheet = context.Book.Worksheets.Item[sheetName];
                range = worksheet.Range[rangeAddress];
                range.Value2 = value;
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref worksheet);
            }
        });

    private static async Task<string> CreateSessionAsync(ExcelMcpService service, string workbookPath)
    {
        var create = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new { filePath = workbookPath, show = ShowExcel }, ServiceProtocol.JsonOptions)
        });
        Assert.True(create.Success, create.ErrorMessage);
        return GetRequiredString(create.Result, "sessionId");
    }

    private static async Task<JsonElement> GetSingleCellAsync(ExcelMcpService service, string sessionId)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.get-values",
            SessionId = sessionId,
            Args = "{\"sheetName\":\"Sheet1\",\"rangeAddress\":\"A1\"}"
        });
        Assert.True(response.Success, response.ErrorMessage);

        using var document = JsonDocument.Parse(response.Result!);
        return document.RootElement.GetProperty("values")[0][0].Clone();
    }

    private static string GetRequiredString(string? json, string propertyName)
    {
        using var document = JsonDocument.Parse(json!);
        return document.RootElement.GetProperty(propertyName).GetString()
            ?? throw new InvalidOperationException($"{propertyName} was null.");
    }

    private static bool GetRequiredBoolean(string? json, string propertyName)
    {
        using var document = JsonDocument.Parse(json!);
        return document.RootElement.GetProperty(propertyName).GetBoolean();
    }

    private static bool ShowExcel =>
        string.Equals(
            Environment.GetEnvironmentVariable("EXCELMCP_TEST_EXCEL_VISIBLE"),
            "true",
            StringComparison.OrdinalIgnoreCase);

    private static void PauseIfVisible()
    {
        if (ShowExcel)
        {
            Thread.Sleep(TimeSpan.FromSeconds(3));
        }
    }

}
