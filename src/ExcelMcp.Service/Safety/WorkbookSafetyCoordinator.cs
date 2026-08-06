// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Generated;

namespace Sbroenne.ExcelMcp.Service.Safety;

/// <summary>
/// Coordinates review, execution, and verification at the shared service-dispatch seam.
/// </summary>
internal sealed class WorkbookSafetyCoordinator : IDisposable
{
    private static readonly TimeSpan ReviewLifetime = TimeSpan.FromMinutes(5);
    private static readonly string[] HighRiskWarnings = ["This operation has a high recovery risk; create a checkpoint before execution."];
    private static readonly string[] StandardWarnings = ["This operation may replace or restructure workbook content."];
    private readonly ConcurrentDictionary<string, SessionSafetyConfiguration> _configurations = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, ReviewAuthorization> _activeReviews = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, string> _terminalReviews = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, SemaphoreSlim> _sessionGates = new(StringComparer.Ordinal);
    private static readonly AsyncLocal<string?> SuppressRequiredCheckpointSession = new();
    private readonly DurableSafetyStore _store;
    private bool _disposed;

    public WorkbookSafetyCoordinator(string? stateRoot = null)
    {
        _store = new DurableSafetyStore(ResolveStateRoot(stateRoot));
    }

    public ServiceResponse Configure(string sessionId, string workbookPath, string? argsJson)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        SessionSafetyConfiguration? configuration;
        try
        {
            if (string.IsNullOrWhiteSpace(argsJson))
            {
                configuration = null;
            }
            else
            {
                using var document = JsonDocument.Parse(argsJson);
                var abnormalShutdownPolicySpecified = document.RootElement.ValueKind == JsonValueKind.Object &&
                    document.RootElement.EnumerateObject().Any(property =>
                        string.Equals(property.Name, "abnormalShutdownPolicy", StringComparison.OrdinalIgnoreCase));
                configuration = JsonSerializer.Deserialize<SessionSafetyConfiguration>(argsJson, ServiceProtocol.JsonOptions)
                    ?.NormalizeAbnormalShutdownPolicy(abnormalShutdownPolicySpecified);
            }
        }
        catch (JsonException ex)
        {
            return Error("InvalidSafetyConfiguration", $"Invalid safety configuration: {ex.Message}");
        }

        if (configuration is null)
        {
            return Error("InvalidSafetyConfiguration", "Safety configuration is required.");
        }

        _configurations[sessionId] = configuration;
        return new ServiceResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(new
            {
                success = true,
                sessionId,
                workbook = Path.GetFileName(workbookPath),
                configuration
            }, ServiceProtocol.JsonOptions)
        };
    }

    /// <summary>
    /// Returns the effective safety policy for a session. Workflow execution uses
    /// this to resolve plan-level checkpoint inheritance before dispatching steps.
    /// </summary>
    internal SessionSafetyConfiguration GetConfiguration(string sessionId) =>
        _configurations.TryGetValue(sessionId, out var configured)
            ? configured
            : SessionSafetyConfiguration.Default;

    /// <summary>Suppresses per-operation required checkpoints after a workflow has created its shared checkpoint.</summary>
    internal static IDisposable SuppressRequiredCheckpoints(string sessionId)
    {
        string? previous = SuppressRequiredCheckpointSession.Value;
        SuppressRequiredCheckpointSession.Value = sessionId;
        return new DelegateDisposable(() => SuppressRequiredCheckpointSession.Value = previous);
    }

    public ServiceResponse Execute(
        ServiceRequest request,
        IExcelBatch batch,
        Func<ServiceResponse> execute)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ArgumentNullException.ThrowIfNull(request);
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(execute);

        var descriptor = ServiceRegistry.GetSafetyDescriptor(request.Command);
        if (!descriptor.IsMutation)
        {
            return execute();
        }

        var sessionId = request.SessionId!;
        var configuration = _configurations.TryGetValue(sessionId, out var configured)
            ? configured
            : SessionSafetyConfiguration.Default;
        var usesWorkflow = configuration.UsesSafetyWorkflow || request.ReviewOnly ||
            !string.IsNullOrWhiteSpace(request.ReviewId) || request.Checkpoint;

        if (!usesWorkflow)
        {
            return execute();
        }

        var gate = _sessionGates.GetOrAdd(sessionId, static _ => new SemaphoreSlim(1, 1));
        gate.Wait();
        try
        {
            return ExecuteInsideGate(request, batch, execute, descriptor, configuration);
        }
        finally
        {
            gate.Release();
        }
    }

    public bool ShouldDiscardOnAbnormalShutdown(string sessionId) =>
        _configurations.TryGetValue(sessionId, out var configuration) &&
        configuration.AbnormalShutdownPolicy == AbnormalShutdownPolicy.DiscardWithRecoveryEvidence;

    public ServiceResponse GetJournal(string sessionId)
    {
        var operations = _store.GetJournal(sessionId);
        return new ServiceResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(new
            {
                success = true,
                sessionId,
                operations
            }, ServiceProtocol.JsonOptions)
        };
    }

    public ServiceResponse ListRecoveries()
    {
        var recoveries = _store.ListRecoveries();
        return new ServiceResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(new
            {
                success = true,
                recoveries,
                count = recoveries.Count
            }, ServiceProtocol.JsonOptions)
        };
    }

    public bool TryResolveRecovery(string recoveryId, out string? checkpointPath, out string? operationId) =>
        _store.TryResolveRecovery(recoveryId, out checkpointPath, out operationId);

    public void RecordRecovered(string operationId) =>
        TryRecordEvidence(() => _store.Transition(operationId, "recovered"));

    public bool RecordSessionInterruption(string sessionId, string state, string category) =>
        TryRecordEvidence(() => _store.TransitionLatestForSession(sessionId, state, category));

    public int RecordServerShutdown(string sessionId)
    {
        var transitioned = 0;
        TryRecordEvidence(() => transitioned =
            _store.TransitionIncompleteForSession(sessionId, "abortedUnknown", "ServerShutdown"));
        return transitioned;
    }

    public void RemoveSession(string sessionId)
    {
        _configurations.TryRemove(sessionId, out _);

        foreach (var review in _activeReviews.Where(pair => pair.Value.SessionId == sessionId).ToArray())
        {
            if (_activeReviews.TryRemove(review.Key, out _))
            {
                _terminalReviews[review.Key] = "ReviewInvalid";
            }
        }
    }

    private ServiceResponse ExecuteInsideGate(
        ServiceRequest request,
        IExcelBatch batch,
        Func<ServiceResponse> execute,
        Sbroenne.ExcelMcp.Core.Models.CommandSafetyDescriptor descriptor,
        SessionSafetyConfiguration configuration)
    {
        var sessionId = request.SessionId!;
        var normalizedArgs = SafetyFingerprint.NormalizeJson(request.Args);
        var workbookIdentity = GetWorkbookIdentity(batch.WorkbookPath);
        var checkpointRequested = request.Checkpoint ||
            (configuration.CheckpointMode == CheckpointMode.Required &&
             !string.Equals(SuppressRequiredCheckpointSession.Value, sessionId, StringComparison.Ordinal));

        if (request.ReviewOnly)
        {
            var baseline = WorkbookSemanticInspector.Capture(batch, descriptor, request.Args);
            var reviewedAtUtc = DateTime.UtcNow;
            var reviewOperationId = Guid.NewGuid().ToString("N");
            var checkpointReservation = checkpointRequested
                ? _store.AllocateCheckpoint(batch.WorkbookPath)
                : null;
            var review = new ReviewAuthorization(
                NewSecureId(),
                reviewOperationId,
                sessionId,
                request.Command,
                normalizedArgs,
                workbookIdentity,
                checkpointRequested,
                checkpointReservation,
                baseline.Fingerprint,
                reviewedAtUtc,
                reviewedAtUtc.Add(ReviewLifetime),
                baseline.Scope);
            _activeReviews[review.ReviewId] = review;
            if (configuration.JournalMode == JournalMode.On)
            {
                _store.BeginReview(review, descriptor.MutationKind);
            }
            return CreateReviewResponse(batch, descriptor, review);
        }

        ReviewAuthorization? authorization = null;
        SemanticSnapshot? before = null;
        if (!string.IsNullOrWhiteSpace(request.ReviewId))
        {
            var validation = ValidateAndConsumeReview(
                request,
                batch,
                descriptor,
                normalizedArgs,
                workbookIdentity,
                checkpointRequested,
                out authorization,
                out before);
            if (validation is not null)
            {
                return validation;
            }
        }
        else if (configuration.ReviewMode == ReviewMode.Required)
        {
            return Error(
                "ReviewRequired",
                $"Command '{request.Command}' requires review. Run the same request with review_only=true, then retry with its review_id.");
        }

        var shouldPersist = configuration.JournalMode == JournalMode.On || checkpointRequested;
        before ??= configuration.VerificationMode == VerificationMode.On || shouldPersist
            ? WorkbookSemanticInspector.Capture(batch, descriptor, request.Args)
            : null;

        var operationId = authorization?.OperationId ?? Guid.NewGuid().ToString("N");
        if (shouldPersist)
        {
            _store.EnsureOperation(
                operationId,
                sessionId,
                request.Command,
                descriptor.MutationKind,
                workbookIdentity,
                before?.Scope ?? authorization?.Scope ?? SafetyScope.Workbook,
                authorization?.ReviewedAtUtc ?? DateTime.UtcNow,
                request.Args);
        }

        CheckpointCreationResult? checkpoint = null;
        if (checkpointRequested)
        {
            try
            {
                var checkpointReservation = authorization?.CheckpointReservation ??
                    _store.AllocateCheckpoint(batch.WorkbookPath);
                _store.Transition(
                    operationId,
                    "checkpointReserved",
                    checkpoint: new SafetyCheckpointRecord(
                        checkpointReservation.RecoveryId,
                        checkpointReservation.RelativePath,
                        string.Empty,
                        0,
                        false,
                        DateTime.UtcNow));
                checkpoint = WorkbookCheckpointManager.Create(
                    batch,
                    _store,
                    checkpointReservation);
                _store.Transition(
                    operationId,
                    "checkpointCreated",
                    checkpoint: new SafetyCheckpointRecord(
                        checkpoint.RecoveryId,
                        checkpoint.RelativePath,
                        checkpoint.Sha256,
                        checkpoint.Size,
                        checkpoint.CalculationSettled,
                        checkpoint.CreatedAtUtc));
            }
            catch (Exception ex)
            {
                if (shouldPersist)
                {
                    TryTransition(operationId, "failed", "CheckpointFailed");
                }

                return Error(
                    "CheckpointFailed",
                    $"Checkpoint creation failed; the mutation was not run. {ex.Message}");
            }
        }

        if (shouldPersist)
        {
            _store.Transition(operationId, "started");
        }

        var stopwatch = Stopwatch.StartNew();
        ServiceResponse response;
        try
        {
            response = execute();
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            if (shouldPersist)
            {
                var (state, category) = ClassifyInterruptedExecution(ex);
                TryTransition(
                    operationId,
                    state,
                    category,
                    durationMilliseconds: stopwatch.ElapsedMilliseconds);
            }

            throw;
        }

        stopwatch.Stop();
        if (!response.Success)
        {
            if (shouldPersist)
            {
                TryTransition(
                    operationId,
                    "failed",
                    response.ErrorCategory ?? "CommandFailed",
                    durationMilliseconds: stopwatch.ElapsedMilliseconds);
            }
            return response;
        }

        if (shouldPersist)
        {
            if (!TryTransition(operationId, "completed", durationMilliseconds: stopwatch.ElapsedMilliseconds))
            {
                return JournalPersistenceError(
                    operationId,
                    "Mutation completed, but its durable completion evidence could not be written. The workbook may have changed; inspect it before retrying.");
            }
        }

        VerificationReceipt verification;
        if (configuration.VerificationMode == VerificationMode.On && before is not null)
        {
            try
            {
                var after = WorkbookSemanticInspector.CapturePostMutation(batch, descriptor, request.Args);
                verification = WorkbookSemanticInspector.Compare(before, after);
            }
            catch (Exception ex) when (!IsFatalVerificationFailure(ex))
            {
                verification = new VerificationReceipt(
                    "failed",
                    before.Scope,
                    0,
                    before.VerificationFingerprint,
                    string.Empty,
                    $"The mutation completed, but post-mutation inspection failed ({ex.GetType().Name}).");
            }
        }
        else
        {
            verification = new VerificationReceipt(
                "notVerified",
                before?.Scope ?? authorization?.Scope ?? SafetyScope.Workbook,
                0,
                before?.VerificationFingerprint ?? string.Empty,
                string.Empty,
                "Session verification mode is off.");
        }

        if (shouldPersist && configuration.VerificationMode == VerificationMode.On)
        {
            var (state, category) = ClassifyVerificationTransition(verification.Status);
            if (!TryTransition(operationId, state, category, verification: verification))
            {
                return JournalPersistenceError(
                    operationId,
                    "Mutation completed, but its durable verification evidence could not be written. The workbook may have changed; inspect it before retrying.");
            }
        }

        return new ServiceResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(new
            {
                success = true,
                executed = true,
                operationId,
                command = request.Command,
                result = JsonResult.Parse(response.Result),
                checkpoint,
                verification
            }, ServiceProtocol.JsonOptions)
        };
    }

    private ServiceResponse? ValidateAndConsumeReview(
        ServiceRequest request,
        IExcelBatch batch,
        Sbroenne.ExcelMcp.Core.Models.CommandSafetyDescriptor descriptor,
        string normalizedArgs,
        string workbookIdentity,
        bool checkpointRequested,
        out ReviewAuthorization? authorization,
        out SemanticSnapshot? before)
    {
        authorization = null;
        before = null;
        var reviewId = request.ReviewId!;

        if (_terminalReviews.TryGetValue(reviewId, out var terminalCategory))
        {
            return ReviewError(terminalCategory);
        }

        if (!_activeReviews.TryGetValue(reviewId, out var review))
        {
            return ReviewError("ReviewInvalid");
        }

        if (review.ExpiresAtUtc <= DateTime.UtcNow)
        {
            Terminalize(reviewId, "ReviewExpired");
            return ReviewError("ReviewExpired");
        }

        if (!string.Equals(review.SessionId, request.SessionId, StringComparison.Ordinal) ||
            !string.Equals(review.Command, request.Command, StringComparison.Ordinal) ||
            !string.Equals(review.NormalizedArgs, normalizedArgs, StringComparison.Ordinal) ||
            !string.Equals(review.WorkbookIdentity, workbookIdentity, StringComparison.Ordinal) ||
            review.CheckpointRequested != checkpointRequested)
        {
            return ReviewError("ReviewInvalid");
        }

        before = WorkbookSemanticInspector.Capture(batch, descriptor, request.Args);
        if (!string.Equals(review.BaselineFingerprint, before.Fingerprint, StringComparison.Ordinal))
        {
            Terminalize(reviewId, "ReviewStale");
            return ReviewError("ReviewStale");
        }

        if (!_activeReviews.TryRemove(reviewId, out authorization))
        {
            return ReviewError(_terminalReviews.TryGetValue(reviewId, out terminalCategory)
                ? terminalCategory
                : "ReviewConsumed");
        }

        _terminalReviews[reviewId] = "ReviewConsumed";
        return null;
    }

    private static ServiceResponse CreateReviewResponse(
        IExcelBatch batch,
        Sbroenne.ExcelMcp.Core.Models.CommandSafetyDescriptor descriptor,
        ReviewAuthorization review)
    {
        return new ServiceResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(new
            {
                success = true,
                executed = false,
                reviewId = review.ReviewId,
                operationId = review.OperationId,
                willWrite = true,
                workbook = Path.GetFileName(batch.WorkbookPath),
                affected = review.Scope,
                saveDestination = batch.WorkbookPath,
                checkpoint = new
                {
                    requested = review.CheckpointRequested,
                    requiredBeforeWrite = review.CheckpointRequested,
                    destination = review.CheckpointReservation?.AbsolutePath
                },
                warnings = descriptor.RecoveryRisk == "high" ? HighRiskWarnings : StandardWarnings,
                verificationPlan = descriptor.VerificationLevel,
                expiresAtUtc = review.ExpiresAtUtc
            }, ServiceProtocol.JsonOptions)
        };
    }

    private void Terminalize(string reviewId, string category)
    {
        _activeReviews.TryRemove(reviewId, out _);
        _terminalReviews[reviewId] = category;
    }

    private static ServiceResponse ReviewError(string category) => category switch
    {
        "ReviewConsumed" => Error(category, "The review ID has already been consumed; the mutation was not repeated."),
        "ReviewExpired" => Error(category, "The review ID has expired; request a new review plan."),
        "ReviewStale" => Error(category, "Workbook state changed after review; request a new review plan."),
        _ => Error("ReviewInvalid", "The review ID does not authorize this exact session, workbook, command, arguments, and checkpoint policy.")
    };

    private static ServiceResponse Error(string category, string message) => new()
    {
        Success = false,
        ErrorCategory = category,
        ErrorMessage = message
    };

    private sealed class DelegateDisposable(Action action) : IDisposable
    {
        private int _disposed;
        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) == 0)
            {
                action();
            }
        }
    }

    private static string GetWorkbookIdentity(string workbookPath) =>
        SafetyFingerprint.Hash(Path.GetFullPath(workbookPath).ToUpperInvariant());

    private static string NewSecureId() =>
        Convert.ToHexString(RandomNumberGenerator.GetBytes(32)).ToLowerInvariant();

    private static (string State, string Category) ClassifyInterruptedExecution(Exception exception) => exception switch
    {
        TimeoutException => ("abortedUnknown", "Timeout"),
        OperationCanceledException => ("abortedUnknown", "Cancelled"),
        _ => ("failed", exception.GetType().Name)
    };

    internal static (string State, string? Category) ClassifyVerificationTransition(string status) => status switch
    {
        "verified" => ("verified", null),
        "partiallyVerified" => ("partiallyVerified", null),
        "notVerified" => ("notVerified", null),
        _ => ("verificationFailed", "VerificationFailed")
    };

    private bool TryTransition(
        string operationId,
        string state,
        string? category = null,
        SafetyCheckpointRecord? checkpoint = null,
        VerificationReceipt? verification = null,
        long? durationMilliseconds = null) => TryRecordEvidence(() => _store.Transition(
                operationId,
                state,
                category,
                checkpoint,
                verification,
                durationMilliseconds));

    internal static ServiceResponse JournalPersistenceError(string operationId, string message) => new()
    {
        Success = false,
        ErrorCategory = "JournalPersistenceFailed",
        ErrorMessage = message,
        Result = JsonSerializer.Serialize(new
        {
            success = false,
            executed = true,
            operationId,
            outcome = "completed-but-durable-evidence-unavailable"
        }, ServiceProtocol.JsonOptions)
    };

    /// <summary>
    /// Writes diagnostic evidence without allowing a secondary journal failure to
    /// replace the original Excel outcome or prevent poisoned-session cleanup.
    /// </summary>
    internal static bool TryRecordEvidence(Action write)
    {
        try
        {
            write();
            return true;
        }
        catch (Exception ex) when (!IsFatalProcessException(ex))
        {
            Debug.WriteLine($"Safety evidence write failed ({ex.GetType().Name}).");
            return false;
        }
    }

    private static bool IsFatalProcessException(Exception exception) =>
        exception is OutOfMemoryException or StackOverflowException or AccessViolationException;

    private static bool IsFatalVerificationFailure(Exception exception)
    {
        for (var current = exception; current is not null; current = current.InnerException)
        {
            if (current is TimeoutException or OperationCanceledException)
            {
                return true;
            }

            if (current is COMException comException &&
                (comException.HResult == ResiliencePipelines.RPC_S_SERVER_UNAVAILABLE ||
                 comException.HResult == ResiliencePipelines.RPC_E_CALL_FAILED ||
                 comException.HResult == ResiliencePipelines.RPC_E_DISCONNECTED))
            {
                return true;
            }
        }

        return false;
    }

    private static string ResolveStateRoot(string? stateRoot)
    {
        var configured = stateRoot;
        if (string.IsNullOrWhiteSpace(configured))
        {
            configured = Environment.GetEnvironmentVariable("EXCELMCP_STATE_DIR");
        }

        if (string.IsNullOrWhiteSpace(configured))
        {
            configured = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Sbroenne",
                "ExcelMcp",
                "safety");
        }

        return Path.GetFullPath(configured);
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        foreach (var gate in _sessionGates.Values)
        {
            gate.Dispose();
        }

        _sessionGates.Clear();
    }
}
