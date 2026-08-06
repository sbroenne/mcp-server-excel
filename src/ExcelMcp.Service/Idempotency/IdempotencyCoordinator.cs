using System.Collections.Concurrent;
using Sbroenne.ExcelMcp.Generated;
using Sbroenne.ExcelMcp.Service.Safety;

namespace Sbroenne.ExcelMcp.Service.Idempotency;

/// <summary>
/// Keeps a bounded, in-memory receipt ledger for retry-safe session requests.
/// The ledger deliberately distinguishes a known completed receipt from an
/// ambiguous timeout/cancellation, which must never be executed again blindly.
/// </summary>
internal sealed class IdempotencyCoordinator
{
    internal const int MaximumKeyLength = 128;
    private const int DefaultMaximumEntries = 1_024;
    private static readonly TimeSpan DefaultRetention = TimeSpan.FromMinutes(30);
    private static readonly TimeSpan DefaultPendingWaitTimeout = TimeSpan.FromSeconds(30);
    private readonly ConcurrentDictionary<string, Entry> _entries = new(StringComparer.Ordinal);
    private readonly object _capacityGate = new();
    private readonly int _maximumEntries;
    private readonly TimeSpan _retention;
    private readonly TimeSpan _pendingWaitTimeout;
    private readonly TimeProvider _timeProvider;

    public IdempotencyCoordinator(
        int maximumEntries = DefaultMaximumEntries,
        TimeSpan? retention = null,
        TimeProvider? timeProvider = null,
        TimeSpan? pendingWaitTimeout = null)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(maximumEntries, 1);
        _maximumEntries = maximumEntries;
        _retention = retention ?? DefaultRetention;
        _pendingWaitTimeout = pendingWaitTimeout ?? DefaultPendingWaitTimeout;
        ArgumentOutOfRangeException.ThrowIfLessThanOrEqual(_pendingWaitTimeout, TimeSpan.Zero);
        _timeProvider = timeProvider ?? TimeProvider.System;
    }

    public async Task<ServiceResponse> ExecuteAsync(
        ServiceRequest request,
        Func<Task<ServiceResponse>> executeAsync)
    {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentNullException.ThrowIfNull(executeAsync);

        if (string.IsNullOrWhiteSpace(request.IdempotencyKey))
        {
            return await executeAsync().ConfigureAwait(false);
        }

        string key = request.IdempotencyKey.Trim();
        if (!IsValidKey(key))
        {
            return CreateFailure(
                request,
                "InvalidIdempotencyKey",
                $"idempotencyKey must be 1-{MaximumKeyLength} characters using only letters, numbers, '.', '_', ':', or '-'.");
        }

        if (string.IsNullOrWhiteSpace(request.SessionId))
        {
            return CreateFailure(
                request,
                "InvalidIdempotencyScope",
                "idempotencyKey requires a session-scoped request.");
        }

        if (request.ReviewOnly)
        {
            return CreateFailure(
                request,
                "InvalidIdempotencyScope",
                "idempotencyKey applies to execution requests, not review-only requests.");
        }

        var descriptor = ServiceRegistry.GetSafetyDescriptor(request.Command);
        bool isWorkflowPlan = string.Equals(request.Command.Trim(), "workflow.execute-plan", StringComparison.OrdinalIgnoreCase);
        if (!isWorkflowPlan && (!descriptor.ExplicitlyClassified || !descriptor.IsMutation))
        {
            return CreateFailure(
                request,
                "InvalidIdempotencyScope",
                "idempotencyKey is supported only for explicitly classified mutation commands.");
        }

        string scope = request.SessionId.Trim();
        string fingerprint = CreateFingerprint(request, scope);
        Entry entry;
        bool ownsEntry;
        lock (_capacityGate)
        {
            TrimExpiredAndOverflowEntries();
            if (_entries.TryGetValue(key, out entry!))
            {
                ownsEntry = false;
            }
            else
            {
                if (_entries.Count >= _maximumEntries)
                {
                    return CreateFailure(
                        request,
                        "IdempotencyCapacityExceeded",
                        "The idempotency receipt ledger is at capacity. Retry after an existing request completes or use a new session.");
                }

                entry = new Entry(scope, fingerprint);
                _entries[key] = entry;
                ownsEntry = true;
            }
        }

        if (!ownsEntry)
        {
            Task<ServiceResponse>? pending = null;
            ServiceResponse? replay = null;
            string? errorCategory = null;
            string? errorMessage = null;

            lock (entry.Gate)
            {
                if (!string.Equals(entry.Scope, scope, StringComparison.Ordinal))
                {
                    errorCategory = "IdempotencyScopeConflict";
                    errorMessage = "The idempotency key is already bound to a different session.";
                }
                else if (!string.Equals(entry.Fingerprint, fingerprint, StringComparison.Ordinal))
                {
                    errorCategory = "IdempotencyConflict";
                    errorMessage = "The idempotency key is already bound to different command arguments or execution options.";
                }
                else
                {
                    switch (entry.State)
                    {
                        case EntryState.Pending:
                            pending = entry.Completion.Task;
                            break;
                        case EntryState.Completed:
                            replay = entry.Response;
                            break;
                        case EntryState.Unknown:
                            errorCategory = "IdempotencyUnknownOutcome";
                            errorMessage = "The original request ended with an ambiguous outcome and will not be replayed automatically. Reconcile the workbook state before issuing a new key.";
                            break;
                    }
                }
            }

            if (pending != null)
            {
                try
                {
                    return await pending.WaitAsync(_pendingWaitTimeout).ConfigureAwait(false);
                }
                catch (TimeoutException)
                {
                    return CreateFailure(
                        request,
                        "IdempotencyInProgress",
                        "The original request is still in progress. Reconcile its status before retrying.");
                }
            }

            if (replay != null)
            {
                return replay;
            }

            var failure = CreateFailure(request, errorCategory!, errorMessage!);
            if (entry.State == EntryState.Unknown && entry.Response?.Result is not null)
            {
                // Preserve the compact plan receipt as reconciliation evidence while
                // still refusing to dispatch the mutation again.
                failure = new ServiceResponse
                {
                    Success = failure.Success,
                    Command = failure.Command,
                    SessionId = failure.SessionId,
                    ErrorCategory = failure.ErrorCategory,
                    ErrorMessage = failure.ErrorMessage,
                    Result = entry.Response.Result
                };
            }

            return failure;
        }

        try
        {
            var response = await executeAsync().ConfigureAwait(false);
            lock (entry.Gate)
            {
                entry.Response = response;
                entry.State = IsAmbiguousOutcome(response)
                    ? EntryState.Unknown
                    : EntryState.Completed;
                entry.TerminalAt = _timeProvider.GetUtcNow();
            }

            if (IsKnownNotExecuted(response))
            {
                // No mutation was dispatched, so this key is safe to use for a real
                // retry. Existing concurrent waiters still receive the exact failure.
                _entries.TryRemove(new KeyValuePair<string, Entry>(key, entry));
            }

            entry.Completion.TrySetResult(response);
            TrimExpiredAndOverflowEntries();
            return response;
        }
        catch (Exception ex)
        {
            _entries.TryRemove(new KeyValuePair<string, Entry>(key, entry));
            entry.Completion.TrySetException(ex);
            throw;
        }
    }

    public void RemoveSession(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return;
        }

        foreach (var pair in _entries)
        {
            if (string.Equals(pair.Value.Scope, sessionId, StringComparison.Ordinal))
            {
                _entries.TryRemove(pair);
            }
        }
    }

    public void Clear() => _entries.Clear();

    private static string CreateFingerprint(ServiceRequest request, string scope) =>
        SafetyFingerprint.Hash(
            scope,
            request.Command.Trim().ToLowerInvariant(),
            SafetyFingerprint.NormalizeJson(request.Args),
            request.ReviewId,
            request.Checkpoint ? "checkpoint" : "no-checkpoint");

    private static bool IsAmbiguousOutcome(ServiceResponse response) =>
        response.ErrorCategory is "Timeout" or "Cancelled" or "ExcelProcessDied" or "UnknownOutcome" or
            "AbortedUnknown" or "IdempotencyUnknownOutcome" or "JournalPersistenceFailed";

    private static bool IsKnownNotExecuted(ServiceResponse response) =>
        response.ErrorCategory is "TimeoutBeforeExecution" or "CancelledBeforeExecution" or "CheckpointFailed" or
            "PlanNotExecuted" or "PlanSafetyConflict" or "PlanOptionConflict" or "PlanReviewUnavailable";

    private static bool IsValidKey(string key)
    {
        if (key.Length is < 1 or > MaximumKeyLength)
        {
            return false;
        }

        foreach (char character in key)
        {
            if (!char.IsAsciiLetterOrDigit(character) && character is not ('.' or '_' or ':' or '-'))
            {
                return false;
            }
        }

        return true;
    }

    private static ServiceResponse CreateFailure(
        ServiceRequest request,
        string errorCategory,
        string errorMessage)
    {
        return new ServiceResponse
        {
            Success = false,
            Command = request.Command,
            SessionId = request.SessionId,
            ErrorCategory = errorCategory,
            ErrorMessage = errorMessage
        };
    }

    private void TrimExpiredAndOverflowEntries()
    {
        var cutoff = _timeProvider.GetUtcNow() - _retention;
        foreach (var pair in _entries)
        {
            var terminalAt = pair.Value.TerminalAt;
            if (terminalAt.HasValue && terminalAt.Value <= cutoff)
            {
                _entries.TryRemove(pair);
            }
        }

        int overflow = _entries.Count - _maximumEntries;
        if (overflow <= 0)
        {
            return;
        }

        foreach (var pair in _entries
                     .Where(pair => pair.Value.TerminalAt.HasValue)
                     .OrderBy(pair => pair.Value.TerminalAt)
                     .Take(overflow))
        {
            _entries.TryRemove(pair);
        }
    }

    private sealed class Entry(string scope, string fingerprint)
    {
        public object Gate { get; } = new();
        public string Scope { get; } = scope;
        public string Fingerprint { get; } = fingerprint;
        public TaskCompletionSource<ServiceResponse> Completion { get; } =
            new(TaskCreationOptions.RunContinuationsAsynchronously);
        public EntryState State { get; set; } = EntryState.Pending;
        public ServiceResponse? Response { get; set; }
        public DateTimeOffset? TerminalAt { get; set; }
    }

    private enum EntryState
    {
        Pending,
        Completed,
        Unknown
    }
}
