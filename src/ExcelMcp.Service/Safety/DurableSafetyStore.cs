// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Security.Cryptography;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.Service.Safety;

/// <summary>
/// Persists sanitized operation state and recovery references beneath a private local root.
/// </summary>
internal sealed class DurableSafetyStore
{
    private static readonly HashSet<string> TerminalStates = new(StringComparer.Ordinal)
    {
        "verified",
        "partiallyVerified",
        "notVerified",
        "verificationFailed",
        "completed",
        "failed",
        "recovered",
        "abortedUnknown",
        "excelProcessDied"
    };
    private readonly object _sync = new();
    private readonly string _root;
    private readonly string _journalDirectory;
    private readonly string _checkpointDirectory;
    private readonly Dictionary<string, SafetyOperationRecord> _operations = new(StringComparer.Ordinal);

    public DurableSafetyStore(string root)
    {
        _root = SafetyStatePathPolicy.PrepareRoot(root);
        _journalDirectory = Path.Combine(_root, "journal");
        _checkpointDirectory = Path.Combine(_root, "checkpoints");
        SafetyStatePathPolicy.EnsureSafePath(_root, _journalDirectory);
        SafetyStatePathPolicy.EnsureSafePath(_root, _checkpointDirectory);
        Directory.CreateDirectory(_journalDirectory);
        Directory.CreateDirectory(_checkpointDirectory);
        SafetyStatePathPolicy.EnsureSafePath(_root, _journalDirectory);
        SafetyStatePathPolicy.EnsureSafePath(_root, _checkpointDirectory);
        LoadExistingOperations();
    }

    public void BeginReview(
        ReviewAuthorization review,
        string mutationKind)
    {
        lock (_sync)
        {
            if (_operations.ContainsKey(review.OperationId))
            {
                return;
            }

            var operation = new SafetyOperationRecord
            {
                OperationId = review.OperationId,
                SessionId = review.SessionId,
                Command = review.Command,
                MutationKind = mutationKind,
                WorkbookIdentity = CreateOpaqueWorkbookReference(review.OperationId, review.WorkbookIdentity),
                Affected = SafetyScopeSummary.From(review.Scope),
                ArgumentSummary = SafetyArgumentSummary.FromJson(review.NormalizedArgs),
                CreatedAtUtc = review.ReviewedAtUtc,
                Transitions = [new SafetyTransition("reviewed", review.ReviewedAtUtc)]
            };
            _operations[operation.OperationId] = operation;
            try { Save(operation); }
            catch { _operations.Remove(operation.OperationId); throw; }
        }
    }

    public void EnsureOperation(
        string operationId,
        string sessionId,
        string command,
        string mutationKind,
        string workbookIdentity,
        SafetyScope scope,
        DateTime createdAtUtc,
        string? argsJson = null)
    {
        lock (_sync)
        {
            if (_operations.ContainsKey(operationId))
            {
                return;
            }

            var operation = new SafetyOperationRecord
            {
                OperationId = operationId,
                SessionId = sessionId,
                Command = command,
                MutationKind = mutationKind,
                WorkbookIdentity = CreateOpaqueWorkbookReference(operationId, workbookIdentity),
                Affected = SafetyScopeSummary.From(scope),
                ArgumentSummary = SafetyArgumentSummary.FromJson(argsJson),
                CreatedAtUtc = createdAtUtc
            };
            _operations[operationId] = operation;
            try { Save(operation); }
            catch { _operations.Remove(operationId); throw; }
        }
    }

    public void Transition(
        string operationId,
        string state,
        string? category = null,
        SafetyCheckpointRecord? checkpoint = null,
        VerificationReceipt? verification = null,
        long? durationMilliseconds = null)
    {
        lock (_sync)
        {
            if (!_operations.TryGetValue(operationId, out var operation))
            {
                throw new InvalidOperationException($"Safety operation '{operationId}' was not initialized.");
            }

            var previous = Clone(operation);
            operation.Transitions.Add(new SafetyTransition(state, DateTime.UtcNow, category));
            operation.Checkpoint = checkpoint ?? operation.Checkpoint;
            operation.Verification = verification is null
                ? operation.Verification
                : VerificationSummary.From(verification);
            operation.OutcomeCategory = category ?? operation.OutcomeCategory;
            operation.DurationMilliseconds = durationMilliseconds ?? operation.DurationMilliseconds;
            try { Save(operation); }
            catch { Restore(operation, previous); throw; }
        }
    }

    public bool TransitionLatestForSession(
        string sessionId,
        string state,
        string category)
    {
        lock (_sync)
        {
            var operation = _operations.Values
                .Where(candidate => string.Equals(candidate.SessionId, sessionId, StringComparison.Ordinal))
                .OrderByDescending(candidate => candidate.CreatedAtUtc)
                .FirstOrDefault();
            if (operation is null)
            {
                return false;
            }

            var latest = operation.Transitions.LastOrDefault();
            if (string.Equals(latest?.State, state, StringComparison.Ordinal) &&
                string.Equals(latest?.Category, category, StringComparison.Ordinal))
            {
                return true;
            }

            var previous = Clone(operation);
            operation.Transitions.Add(new SafetyTransition(state, DateTime.UtcNow, category));
            operation.OutcomeCategory = category;
            try { Save(operation); }
            catch { Restore(operation, previous); throw; }
            return true;
        }
    }

    /// <summary>
    /// Records a server interruption for every durable operation in the session that
    /// has not already reached a terminal outcome. Completed operations remain intact.
    /// </summary>
    public int TransitionIncompleteForSession(
        string sessionId,
        string state,
        string category)
    {
        lock (_sync)
        {
            var transitioned = 0;
            foreach (var operation in _operations.Values.Where(candidate =>
                         string.Equals(candidate.SessionId, sessionId, StringComparison.Ordinal)))
            {
                var latest = operation.Transitions.LastOrDefault();
                if (latest is not null && TerminalStates.Contains(latest.State))
                {
                    continue;
                }

                if (string.Equals(latest?.State, state, StringComparison.Ordinal) &&
                    string.Equals(latest?.Category, category, StringComparison.Ordinal))
                {
                    continue;
                }

                var previous = Clone(operation);
                operation.Transitions.Add(new SafetyTransition(state, DateTime.UtcNow, category));
                operation.OutcomeCategory = category;
                try
                {
                    Save(operation);
                    transitioned++;
                }
                catch (Exception ex) when (IsUnavailableCheckpointException(ex))
                {
                    Restore(operation, previous);
                }
            }

            return transitioned;
        }
    }

    public IReadOnlyList<SafetyOperationRecord> GetJournal(string sessionId)
    {
        lock (_sync)
        {
            return _operations.Values
                .Where(operation => string.Equals(operation.SessionId, sessionId, StringComparison.Ordinal))
                .OrderBy(operation => operation.CreatedAtUtc)
                .Select(Clone)
                .ToArray();
        }
    }

    public IReadOnlyList<object> ListRecoveries()
    {
        lock (_sync)
        {
            return _operations.Values
                .Where(operation => operation.Checkpoint is not null)
                .OrderByDescending(operation => operation.Checkpoint!.CreatedAtUtc)
                .Select(operation =>
                {
                    var checkpoint = operation.Checkpoint!;
                    return (object)new
                    {
                        recoveryId = checkpoint.RecoveryId,
                        operationId = operation.OperationId,
                        workbookReference = operation.WorkbookIdentity,
                        command = operation.Command,
                        checkpoint.CreatedAtUtc,
                        status = operation.Transitions.LastOrDefault()?.State ?? "unknown",
                        available = IsCheckpointAvailable(checkpoint)
                    };
                })
                .ToArray();
        }
    }

    public bool TryResolveRecovery(string recoveryId, out string? checkpointPath, out string? operationId)
    {
        lock (_sync)
        {
            var operation = _operations.Values.FirstOrDefault(candidate =>
                string.Equals(candidate.Checkpoint?.RecoveryId, recoveryId, StringComparison.Ordinal));
            if (operation?.Checkpoint is null)
            {
                checkpointPath = null;
                operationId = null;
                return false;
            }

            try
            {
                var path = ResolveRelativePath(operation.Checkpoint.RelativePath);
                if (!File.Exists(path))
                {
                    checkpointPath = null;
                    operationId = operation.OperationId;
                    return false;
                }

                var fileInfo = new FileInfo(path);
                if (fileInfo.Length <= 0 ||
                    fileInfo.Length != operation.Checkpoint.Size ||
                    !string.Equals(ComputeFileHash(path), operation.Checkpoint.Sha256, StringComparison.Ordinal))
                {
                    checkpointPath = null;
                    operationId = operation.OperationId;
                    return false;
                }

                checkpointPath = path;
                operationId = operation.OperationId;
                return true;
            }
            catch (Exception ex) when (IsUnavailableCheckpointException(ex))
            {
                checkpointPath = null;
                operationId = operation.OperationId;
                return false;
            }
        }
    }

    public CheckpointReservation AllocateCheckpoint(
        string workbookPath)
    {
        var recoveryId = Guid.NewGuid().ToString("N");
        var extension = Path.GetExtension(workbookPath).ToLowerInvariant();
        if (extension is not ".xlsx" and not ".xlsm" and not ".xls")
        {
            throw new InvalidOperationException($"Cannot checkpoint workbook extension '{extension}'.");
        }

        var fileName = $"{recoveryId}{extension}";
        var relativePath = Path.Combine("checkpoints", recoveryId, fileName);
        var absolutePath = ResolveRelativePath(relativePath);
        var checkpointParent = Path.GetDirectoryName(absolutePath)!;
        SafetyStatePathPolicy.EnsureSafePath(_root, checkpointParent);
        Directory.CreateDirectory(checkpointParent);
        SafetyStatePathPolicy.EnsureSafePath(_root, absolutePath);
        if (File.Exists(absolutePath))
        {
            throw new IOException("Allocated checkpoint path already exists.");
        }

        return new CheckpointReservation(recoveryId, absolutePath, relativePath);
    }

    public static string ComputeFileHash(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }

    public void EnsureSafeCheckpointPath(string path) =>
        SafetyStatePathPolicy.EnsureSafePath(_root, path);

    private bool IsCheckpointAvailable(SafetyCheckpointRecord checkpoint)
    {
        if (checkpoint.Size <= 0 || string.IsNullOrWhiteSpace(checkpoint.Sha256))
        {
            return false;
        }

        try
        {
            var path = ResolveRelativePath(checkpoint.RelativePath);
            return File.Exists(path) &&
                new FileInfo(path).Length == checkpoint.Size &&
                string.Equals(ComputeFileHash(path), checkpoint.Sha256, StringComparison.Ordinal);
        }
        catch (Exception ex) when (IsUnavailableCheckpointException(ex))
        {
            return false;
        }
    }

    private static bool IsUnavailableCheckpointException(Exception exception) =>
        exception is IOException or UnauthorizedAccessException or InvalidOperationException or ArgumentException or NotSupportedException;

    private void LoadExistingOperations()
    {
        var loaded = new List<SafetyOperationRecord>();
        foreach (var path in Directory
                     .EnumerateFiles(_journalDirectory, "*.json", SearchOption.TopDirectoryOnly)
                     .Order(StringComparer.Ordinal))
        {
            try
            {
                SafetyStatePathPolicy.EnsureSafePath(_root, path);
                var operation = JsonSerializer.Deserialize<SafetyOperationRecord>(
                    File.ReadAllText(path),
                    ServiceProtocol.JsonOptions);
                if (operation is null || string.IsNullOrWhiteSpace(operation.OperationId))
                {
                    throw new InvalidDataException(
                        $"Safety journal '{Path.GetFileName(path)}' is corrupt because it has no operation identity.");
                }

                if (!string.Equals(
                        Path.GetFileNameWithoutExtension(path),
                        operation.OperationId,
                        StringComparison.Ordinal))
                {
                    throw new InvalidDataException(
                        $"Safety journal '{Path.GetFileName(path)}' is corrupt because its operation identity does not match its evidence filename.");
                }

                loaded.Add(operation);
            }
            catch (JsonException ex)
            {
                throw new InvalidDataException(
                    $"Safety journal '{Path.GetFileName(path)}' is corrupt and cannot be trusted. Repair or quarantine the evidence file before starting the service.",
                    ex);
            }
            catch (InvalidDataException)
            {
                throw;
            }
            catch (IOException ex)
            {
                throw new InvalidDataException(
                    $"Safety journal '{Path.GetFileName(path)}' could not be read completely. The service is failing closed to avoid loading partial evidence.",
                    ex);
            }
        }

        foreach (var operation in loaded)
        {
            if (!_operations.TryAdd(operation.OperationId, operation))
            {
                throw new InvalidDataException(
                    $"Safety journal contains duplicate operation identity '{operation.OperationId}'.");
            }

            FinalizePendingCheckpoint(operation);
        }
    }

    private void FinalizePendingCheckpoint(SafetyOperationRecord operation)
    {
        var checkpoint = operation.Checkpoint;
        if (checkpoint is null ||
            !string.IsNullOrEmpty(checkpoint.Sha256) ||
            checkpoint.Size != 0)
        {
            return;
        }

        var previous = Clone(operation);
        try
        {
            var path = ResolveRelativePath(checkpoint.RelativePath);
            var pendingPath = WorkbookCheckpointManager.GetPendingCheckpointPath(path);
            var readyMarkerPath = WorkbookCheckpointManager.GetReadyMarkerPath(path);
            EnsureSafeCheckpointPath(pendingPath);
            EnsureSafeCheckpointPath(readyMarkerPath);

            if (!File.Exists(path))
            {
                // A staged copy is trusted only when the durable ready marker
                // matches the staged bytes. Without that marker, the process may
                // have died before the checkpoint flush completed.
                if (!File.Exists(pendingPath) ||
                    !File.Exists(readyMarkerPath) ||
                    !WorkbookCheckpointManager.TryReadReadyMarker(readyMarkerPath, out var marker) ||
                    marker is null)
                {
                    return;
                }

                var pendingInfo = new FileInfo(pendingPath);
                if (pendingInfo.Length <= 0 || pendingInfo.Length != marker.Size ||
                    !string.Equals(ComputeFileHash(pendingPath), marker.Sha256, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                DurableFileWriter.PublishFlushedFileAtomically(pendingPath, path);
            }

            var fileInfo = new FileInfo(path);
            if (fileInfo.Length <= 0)
            {
                return;
            }

            var sizeBeforeHash = fileInfo.Length;
            var hash = ComputeFileHash(path);
            fileInfo.Refresh();
            if (fileInfo.Length <= 0 || fileInfo.Length != sizeBeforeHash)
            {
                return;
            }

            operation.Checkpoint = checkpoint with
            {
                Sha256 = hash,
                Size = fileInfo.Length
            };
            operation.Transitions.Add(new SafetyTransition("checkpointCreated", DateTime.UtcNow));
            Save(operation);
            TryDeleteReadyMarker(readyMarkerPath);
        }
        catch (Exception ex) when (IsUnavailableCheckpointException(ex))
        {
            // Preserve both durable and in-memory state when publication or journal
            // persistence is unavailable; a later restart can retry safely.
            Restore(operation, previous);
        }
    }

    private static void TryDeleteReadyMarker(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (IOException)
        {
            // Marker cleanup is best effort after the final checkpoint is trusted.
        }
        catch (UnauthorizedAccessException)
        {
            // Marker cleanup is best effort after the final checkpoint is trusted.
        }
    }

    private void Save(SafetyOperationRecord operation)
    {
        var path = Path.Combine(_journalDirectory, $"{operation.OperationId}.json");
        SafetyStatePathPolicy.EnsureSafePath(_root, path);
        var json = JsonSerializer.Serialize(operation, ServiceProtocol.JsonOptions);
        DurableFileWriter.WriteUtf8Atomically(path, json);
    }

    private string ResolveRelativePath(string relativePath)
    {
        var fullPath = Path.GetFullPath(Path.Combine(_root, relativePath));
        var relative = Path.GetRelativePath(_root, fullPath);
        if (relative.StartsWith("..", StringComparison.Ordinal) || Path.IsPathRooted(relative))
        {
            throw new InvalidOperationException("Recovery path escaped the configured safety-state root.");
        }

        SafetyStatePathPolicy.EnsureSafePath(_root, fullPath);
        return fullPath;
    }

    private static string CreateOpaqueWorkbookReference(string operationId, string workbookIdentity) =>
        SafetyFingerprint.Hash("journal-workbook", operationId, workbookIdentity);

    private static SafetyOperationRecord Clone(SafetyOperationRecord operation)
    {
        return new SafetyOperationRecord
        {
            OperationId = operation.OperationId,
            SessionId = operation.SessionId,
            Command = operation.Command,
            MutationKind = operation.MutationKind,
            WorkbookIdentity = operation.WorkbookIdentity,
            Affected = operation.Affected,
            ArgumentSummary = operation.ArgumentSummary,
            CreatedAtUtc = operation.CreatedAtUtc,
            Transitions = [.. operation.Transitions],
            Checkpoint = operation.Checkpoint,
            Verification = operation.Verification,
            OutcomeCategory = operation.OutcomeCategory,
            DurationMilliseconds = operation.DurationMilliseconds
        };
    }

    private static void Restore(SafetyOperationRecord target, SafetyOperationRecord snapshot)
    {
        target.Transitions.Clear();
        target.Transitions.AddRange(snapshot.Transitions);
        target.Checkpoint = snapshot.Checkpoint;
        target.Verification = snapshot.Verification;
        target.OutcomeCategory = snapshot.OutcomeCategory;
        target.DurationMilliseconds = snapshot.DurationMilliseconds;
    }
}
