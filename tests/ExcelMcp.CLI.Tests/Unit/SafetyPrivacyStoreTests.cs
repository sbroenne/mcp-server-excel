using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Service.Safety;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "SafetyPrivacy")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class SafetyPrivacyStoreTests : IDisposable
{
    private readonly string _stateRoot = Path.Combine(
        Path.GetTempPath(),
        $"excelmcp-safety-privacy-{Guid.NewGuid():N}");

    [Fact]
    public void CorruptJournal_FailsClosedWithEvidenceFileContext()
    {
        var journalDirectory = Path.Combine(_stateRoot, "journal");
        Directory.CreateDirectory(journalDirectory);
        const string fileName = "truncated-operation.json";
        File.WriteAllText(Path.Combine(journalDirectory, fileName), "{\"operationId\":");

        var failure = Assert.Throws<InvalidDataException>(() =>
            new DurableSafetyStore(_stateRoot));

        Assert.Contains(fileName, failure.Message, StringComparison.Ordinal);
        Assert.Contains("corrupt", failure.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void LegacyJournalWithoutArgumentSummary_LoadsWithEmptyPrivacySafeSummary()
    {
        var journalDirectory = Path.Combine(_stateRoot, "journal");
        Directory.CreateDirectory(journalDirectory);
        File.WriteAllText(
            Path.Combine(journalDirectory, "legacy-operation.json"),
            """
            {
              "operationId": "legacy-operation",
              "sessionId": "legacy-session",
              "command": "range.set-values",
              "mutationKind": "values",
              "workbookIdentity": "opaque-workbook",
              "affected": {
                "sheetCount": 1,
                "rangeCount": 1,
                "objectCount": 0
              },
              "createdAtUtc": "2026-08-05T12:00:00Z",
              "transitions": []
            }
            """);

        var restarted = new DurableSafetyStore(_stateRoot);

        var operation = Assert.Single(restarted.GetJournal("legacy-session"));
        Assert.Equal(0, operation.ArgumentSummary.ParameterCount);
        Assert.Equal(0, operation.ArgumentSummary.StringCount);
        Assert.Equal(0, operation.ArgumentSummary.NumberCount);
        Assert.Equal(0, operation.ArgumentSummary.BooleanCount);
        Assert.Equal(0, operation.ArgumentSummary.ObjectCount);
        Assert.Equal(0, operation.ArgumentSummary.ArrayCount);
        Assert.Equal(0, operation.ArgumentSummary.NullCount);
    }

    [Fact]
    public void DurableMetadata_StoresOnlyOpaqueWorkbookAndScopeSummaries()
    {
        const string workbookName = "AcmeSecretWorkbook";
        const string workbookIdentity = "RawWorkbookIdentitySentinel";
        const string sheetName = "AcmeSecretSheet";
        const string rangeAddress = "AcmeSecretSheet!$A$1";
        const string objectName = "AcmeSecretObject";
        const string beforeFingerprint = "RawBeforeFingerprintSentinel";
        const string afterFingerprint = "RawAfterFingerprintSentinel";
        const string operationId = "0123456789abcdef0123456789abcdef";
        const string sessionId = "privacy-session";

        var store = new DurableSafetyStore(_stateRoot);
        var workbookPath = $@"C:\Clients\{workbookName}.xlsx";
        store.EnsureOperation(
            operationId,
            sessionId,
            "range.set-values",
            "values",
            workbookIdentity,
            new SafetyScope([sheetName], [rangeAddress], [objectName]),
            DateTime.UtcNow);

        var reservation = store.AllocateCheckpoint(workbookPath);
        File.WriteAllBytes(reservation.AbsolutePath, [1, 2, 3, 4]);
        store.Transition(
            operationId,
            "checkpointCreated",
            checkpoint: new SafetyCheckpointRecord(
                reservation.RecoveryId,
                reservation.RelativePath,
                DurableSafetyStore.ComputeFileHash(reservation.AbsolutePath),
                4,
                true,
                DateTime.UtcNow));
        store.Transition(
            operationId,
            "verified",
            verification: new VerificationReceipt(
                "verified",
                new SafetyScope([sheetName], [rangeAddress], []),
                1,
                beforeFingerprint,
                afterFingerprint,
                null));

        var persisted = string.Join(
            Environment.NewLine,
            Directory.EnumerateFiles(_stateRoot, "*.json", SearchOption.AllDirectories)
                .Select(File.ReadAllText));
        var journal = JsonSerializer.Serialize(store.GetJournal(sessionId), ServiceProtocol.JsonOptions);
        var recoveries = JsonSerializer.Serialize(store.ListRecoveries(), ServiceProtocol.JsonOptions);
        var allPortableMetadata = string.Join(Environment.NewLine, persisted, journal, recoveries, reservation.RelativePath);

        foreach (var sensitiveValue in new[]
                 {
                     workbookName,
                     workbookIdentity,
                     sheetName,
                     rangeAddress,
                     objectName,
                     beforeFingerprint,
                     afterFingerprint
                 })
        {
            Assert.DoesNotContain(sensitiveValue, allPortableMetadata, StringComparison.OrdinalIgnoreCase);
        }

        using var journalDocument = JsonDocument.Parse(journal);
        var operation = journalDocument.RootElement[0];
        Assert.Equal(operationId, operation.GetProperty("operationId").GetString());
        Assert.Equal("range.set-values", operation.GetProperty("command").GetString());
        Assert.Equal(1, operation.GetProperty("affected").GetProperty("sheetCount").GetInt32());
        Assert.Equal(1, operation.GetProperty("affected").GetProperty("rangeCount").GetInt32());
        Assert.Equal(1, operation.GetProperty("affected").GetProperty("objectCount").GetInt32());
        Assert.Equal("verified", operation.GetProperty("verification").GetProperty("status").GetString());
    }

    [Fact]
    public void PendingCheckpointWithExistingFile_IsFinalizedAndRecoverableAfterRestart()
    {
        const string operationId = "pending-checkpoint-operation";
        const string sessionId = "pending-checkpoint-session";
        var store = new DurableSafetyStore(_stateRoot);
        store.EnsureOperation(
            operationId,
            sessionId,
            "range.set-values",
            "values",
            "workbook",
            SafetyScope.Workbook,
            DateTime.UtcNow);

        var reservation = store.AllocateCheckpoint(@"C:\Clients\Workbook.xlsx");
        store.Transition(
            operationId,
            "checkpointReserved",
            checkpoint: new SafetyCheckpointRecord(
                reservation.RecoveryId,
                reservation.RelativePath,
                string.Empty,
                0,
                false,
                DateTime.UtcNow));
        File.WriteAllBytes(reservation.AbsolutePath, [1, 2, 3, 4]);

        var restarted = new DurableSafetyStore(_stateRoot);

        Assert.True(restarted.TryResolveRecovery(
            reservation.RecoveryId,
            out var checkpointPath,
            out var recoveredOperationId));
        Assert.Equal(reservation.AbsolutePath, checkpointPath);
        Assert.Equal(operationId, recoveredOperationId);

        var operation = Assert.Single(restarted.GetJournal(sessionId));
        Assert.Equal(4, operation.Checkpoint!.Size);
        Assert.Equal(
            DurableSafetyStore.ComputeFileHash(reservation.AbsolutePath),
            operation.Checkpoint.Sha256);
        Assert.Equal("checkpointCreated", operation.Transitions[^1].State);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PendingCheckpointWithoutNonEmptyFile_IsNotRecoverable(bool createEmptyFile)
    {
        const string operationId = "pending-checkpoint-rejected";
        const string sessionId = "pending-checkpoint-rejected-session";
        var store = new DurableSafetyStore(_stateRoot);
        store.EnsureOperation(
            operationId,
            sessionId,
            "range.set-values",
            "values",
            "workbook",
            SafetyScope.Workbook,
            DateTime.UtcNow);

        var reservation = store.AllocateCheckpoint(@"C:\Clients\Workbook.xlsx");
        store.Transition(
            operationId,
            "checkpointReserved",
            checkpoint: new SafetyCheckpointRecord(
                reservation.RecoveryId,
                reservation.RelativePath,
                string.Empty,
                0,
                false,
                DateTime.UtcNow));
        if (createEmptyFile)
        {
            File.WriteAllBytes(reservation.AbsolutePath, []);
        }

        var restarted = new DurableSafetyStore(_stateRoot);

        Assert.False(restarted.TryResolveRecovery(
            reservation.RecoveryId,
            out var checkpointPath,
            out var recoveredOperationId));
        Assert.Null(checkpointPath);
        Assert.Equal(operationId, recoveredOperationId);
        var operation = Assert.Single(restarted.GetJournal(sessionId));
        Assert.Empty(operation.Checkpoint!.Sha256);
        Assert.Equal(0, operation.Checkpoint.Size);
        Assert.Equal("checkpointReserved", operation.Transitions[^1].State);
    }

    [Fact]
    public void PendingCheckpointWithReadyMarker_IsPublishedAndRecoverableAfterRestart()
    {
        const string operationId = "pending-checkpoint-ready";
        const string sessionId = "pending-checkpoint-ready-session";
        var store = new DurableSafetyStore(_stateRoot);
        store.EnsureOperation(
            operationId,
            sessionId,
            "range.set-values",
            "values",
            "workbook",
            SafetyScope.Workbook,
            DateTime.UtcNow);

        var reservation = store.AllocateCheckpoint(@"C:\Clients\Workbook.xlsx");
        store.Transition(
            operationId,
            "checkpointReserved",
            checkpoint: new SafetyCheckpointRecord(
                reservation.RecoveryId,
                reservation.RelativePath,
                string.Empty,
                0,
                false,
                DateTime.UtcNow));

        var pendingPath = WorkbookCheckpointManager.GetPendingCheckpointPath(reservation.AbsolutePath);
        var bytes = new byte[] { 1, 2, 3, 4, 5 };
        File.WriteAllBytes(pendingPath, bytes);
        DurableFileWriter.FlushExistingFile(pendingPath);
        DurableFileWriter.WriteUtf8Atomically(
            WorkbookCheckpointManager.GetReadyMarkerPath(reservation.AbsolutePath),
            JsonSerializer.Serialize(
                new CheckpointReadyMarker(bytes.Length, DurableSafetyStore.ComputeFileHash(pendingPath)),
                ServiceProtocol.JsonOptions));

        var restarted = new DurableSafetyStore(_stateRoot);

        Assert.True(restarted.TryResolveRecovery(
            reservation.RecoveryId,
            out var checkpointPath,
            out var recoveredOperationId));
        Assert.Equal(reservation.AbsolutePath, checkpointPath);
        Assert.Equal(operationId, recoveredOperationId);
        Assert.True(File.Exists(reservation.AbsolutePath));
        Assert.False(File.Exists(pendingPath));
        Assert.False(File.Exists(WorkbookCheckpointManager.GetReadyMarkerPath(reservation.AbsolutePath)));
        var operation = Assert.Single(restarted.GetJournal(sessionId));
        Assert.Equal(bytes.Length, operation.Checkpoint!.Size);
        Assert.Equal("checkpointCreated", operation.Transitions[^1].State);
    }

    [Fact]
    public void PendingCheckpointWithoutReadyMarker_IsNotTrustedAfterRestart()
    {
        const string operationId = "pending-checkpoint-no-marker";
        const string sessionId = "pending-checkpoint-no-marker-session";
        var store = new DurableSafetyStore(_stateRoot);
        store.EnsureOperation(
            operationId,
            sessionId,
            "range.set-values",
            "values",
            "workbook",
            SafetyScope.Workbook,
            DateTime.UtcNow);

        var reservation = store.AllocateCheckpoint(@"C:\Clients\Workbook.xlsx");
        store.Transition(
            operationId,
            "checkpointReserved",
            checkpoint: new SafetyCheckpointRecord(
                reservation.RecoveryId,
                reservation.RelativePath,
                string.Empty,
                0,
                false,
                DateTime.UtcNow));

        var pendingPath = WorkbookCheckpointManager.GetPendingCheckpointPath(reservation.AbsolutePath);
        File.WriteAllBytes(pendingPath, [1, 2, 3, 4]);

        var restarted = new DurableSafetyStore(_stateRoot);

        Assert.False(restarted.TryResolveRecovery(
            reservation.RecoveryId,
            out var checkpointPath,
            out var recoveredOperationId));
        Assert.Null(checkpointPath);
        Assert.Equal(operationId, recoveredOperationId);
        Assert.True(File.Exists(pendingPath));
        var operation = Assert.Single(restarted.GetJournal(sessionId));
        Assert.Equal(0, operation.Checkpoint!.Size);
        Assert.Equal("checkpointReserved", operation.Transitions[^1].State);
    }

    [Fact]
    public void PendingCheckpointWithMismatchedReadyMarker_IsNotTrustedAfterRestart()
    {
        const string operationId = "pending-checkpoint-bad-marker";
        const string sessionId = "pending-checkpoint-bad-marker-session";
        var store = new DurableSafetyStore(_stateRoot);
        store.EnsureOperation(
            operationId,
            sessionId,
            "range.set-values",
            "values",
            "workbook",
            SafetyScope.Workbook,
            DateTime.UtcNow);

        var reservation = store.AllocateCheckpoint(@"C:\Clients\Workbook.xlsx");
        store.Transition(
            operationId,
            "checkpointReserved",
            checkpoint: new SafetyCheckpointRecord(
                reservation.RecoveryId,
                reservation.RelativePath,
                string.Empty,
                0,
                false,
                DateTime.UtcNow));

        var pendingPath = WorkbookCheckpointManager.GetPendingCheckpointPath(reservation.AbsolutePath);
        File.WriteAllBytes(pendingPath, [1, 2, 3, 4]);
        DurableFileWriter.WriteUtf8Atomically(
            WorkbookCheckpointManager.GetReadyMarkerPath(reservation.AbsolutePath),
            JsonSerializer.Serialize(
                new CheckpointReadyMarker(999, new string('0', 64)),
                ServiceProtocol.JsonOptions));

        var restarted = new DurableSafetyStore(_stateRoot);

        Assert.False(restarted.TryResolveRecovery(reservation.RecoveryId, out _, out var recoveredOperationId));
        Assert.Equal(operationId, recoveredOperationId);
        Assert.True(File.Exists(pendingPath));
        Assert.False(File.Exists(reservation.AbsolutePath));
    }

    [Fact]
    public void CorruptCheckpointPath_IsListedUnavailableAndCannotEscapeRecoveryRoot()
    {
        const string operationId = "corrupt-checkpoint-operation";
        const string sessionId = "corrupt-checkpoint-session";
        const string recoveryId = "corrupt-checkpoint-recovery";
        var journalDirectory = Path.Combine(_stateRoot, "journal");
        Directory.CreateDirectory(journalDirectory);
        var operation = new SafetyOperationRecord
        {
            OperationId = operationId,
            SessionId = sessionId,
            Command = "range.set-values",
            MutationKind = "values",
            WorkbookIdentity = "opaque",
            Affected = new SafetyScopeSummary(0, 1, 0),
            CreatedAtUtc = DateTime.UtcNow,
            Checkpoint = new SafetyCheckpointRecord(
                recoveryId,
                Path.Combine("..", "outside.xlsx"),
                new string('a', 64),
                1,
                false,
                DateTime.UtcNow)
        };
        File.WriteAllText(
            Path.Combine(journalDirectory, $"{operationId}.json"),
            JsonSerializer.Serialize(operation, ServiceProtocol.JsonOptions));

        var restarted = new DurableSafetyStore(_stateRoot);

        using var recoveries = JsonDocument.Parse(
            JsonSerializer.Serialize(restarted.ListRecoveries(), ServiceProtocol.JsonOptions));
        Assert.False(recoveries.RootElement[0].GetProperty("available").GetBoolean());
        Assert.False(restarted.TryResolveRecovery(recoveryId, out var checkpointPath, out var recoveredOperationId));
        Assert.Null(checkpointPath);
        Assert.Equal(operationId, recoveredOperationId);
    }

    [Fact]
    public void SameSizeCheckpointTamper_IsNotAdvertisedAsAvailable()
    {
        const string operationId = "tampered-checkpoint-operation";
        var store = new DurableSafetyStore(_stateRoot);
        store.EnsureOperation(
            operationId,
            "tampered-checkpoint-session",
            "range.set-values",
            "values",
            "workbook",
            SafetyScope.Workbook,
            DateTime.UtcNow);
        var reservation = store.AllocateCheckpoint(@"C:\Clients\Workbook.xlsx");
        File.WriteAllBytes(reservation.AbsolutePath, [1, 2, 3, 4]);
        store.Transition(
            operationId,
            "checkpointCreated",
            checkpoint: new SafetyCheckpointRecord(
                reservation.RecoveryId,
                reservation.RelativePath,
                DurableSafetyStore.ComputeFileHash(reservation.AbsolutePath),
                4,
                true,
                DateTime.UtcNow));

        File.WriteAllBytes(reservation.AbsolutePath, [4, 3, 2, 1]);

        using var recoveries = JsonDocument.Parse(
            JsonSerializer.Serialize(store.ListRecoveries(), ServiceProtocol.JsonOptions));
        Assert.False(recoveries.RootElement[0].GetProperty("available").GetBoolean());
        Assert.False(store.TryResolveRecovery(reservation.RecoveryId, out _, out _));
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (Directory.Exists(_stateRoot))
        {
            Directory.Delete(_stateRoot, recursive: true);
        }

        GC.SuppressFinalize(this);
    }
}
