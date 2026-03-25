using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>
/// OLE Message Filter for handling Excel COM busy/retry scenarios.
/// Automatically retries when Excel returns RPC_E_SERVERCALL_RETRYLATER.
/// </summary>
/// <remarks>
/// This filter intercepts COM calls to Excel and handles transient "server busy" conditions.
/// When Excel is temporarily busy (e.g., showing a dialog), the filter automatically retries
/// after a short delay rather than throwing an exception.
///
/// Register once per STA thread via Register(), revoke on thread shutdown via Revoke().
/// </remarks>
[GeneratedComClass]
public sealed partial class OleMessageFilter : IOleMessageFilter
{
    private static readonly StrategyBasedComWrappers s_comWrappers = new();

    [ThreadStatic]
    private static nint _oldFilterPtr;

    [ThreadStatic]
    private static bool _isRegistered;

    /// <summary>
    /// When true, the filter is in a long-running COM operation (e.g., Power Query refresh).
    /// MessagePending returns WAITDEFPROCESS to dispatch to HandleInComingCall, which rejects
    /// with SERVERCALL_RETRYLATER to trigger the caller's RetryRejectedCall backoff.
    /// </summary>
    [ThreadStatic]
    private static volatile bool _isInLongOperation;

    /// <summary>
    /// Diagnostic counter: total MessagePending calls during the current long operation.
    /// Reset on EnterLongOperation, read on ExitLongOperation.
    /// </summary>
    [ThreadStatic]
    private static long _messagePendingCount;

    /// <summary>
    /// Diagnostic counter: total HandleInComingCall rejections during the current long operation.
    /// </summary>
    [ThreadStatic]
    private static long _handleInComingCallRejections;

    /// <summary>
    /// Timestamp when the current long operation started (for diagnostics).
    /// </summary>
    [ThreadStatic]
    private static long _longOperationStartTimestamp;

    /// <summary>
    /// CancellationToken associated with the current outgoing COM call on this STA thread.
    /// When cancelled, <c>IMessageFilter.MessagePending</c> returns PENDINGMSG_CANCELCALL (0) to abort
    /// the pending outgoing call (e.g., connection.Refresh()) so the STA thread is not orphaned.
    /// Set via <see cref="SetPendingCancellationToken"/> before the COM call; cleared in finally.
    /// </summary>
    [ThreadStatic]
    private static CancellationToken _pendingCancellationToken;

    /// <summary>
    /// Registers the OLE message filter for the current STA thread.
    /// Should be called once per STA thread before making COM calls.
    /// </summary>
    /// <exception cref="InvalidOperationException">Filter already registered on this thread, or registration failed</exception>
    public static void Register()
    {
        if (_isRegistered)
        {
            throw new InvalidOperationException("OLE message filter is already registered on this thread.");
        }

        var newFilter = new OleMessageFilter();
        nint newFilterPtr = s_comWrappers.GetOrCreateComInterfaceForObject(newFilter, CreateComInterfaceFlags.None);

        int result = CoRegisterMessageFilter(newFilterPtr, out _oldFilterPtr);
        if (result != 0)
        {
            throw new InvalidOperationException($"Failed to register OLE message filter. HRESULT: 0x{result:X8}");
        }

        _isRegistered = true;
    }

    /// <summary>
    /// Revokes the OLE message filter and restores the previous filter.
    /// Should be called when STA thread is shutting down.
    /// </summary>
    /// <remarks>
    /// This method is safe to call even if Register() was not called - it will simply return.
    /// This supports cleanup scenarios where the registration status is unknown.
    /// </remarks>
    public static void Revoke()
    {
        if (!_isRegistered)
        {
            // Safe to call without prior Register - just return silently
            return;
        }

        int result = CoRegisterMessageFilter(_oldFilterPtr, out _);
        if (result != 0)
        {
            throw new InvalidOperationException($"Failed to revoke OLE message filter. HRESULT: 0x{result:X8}");
        }

        _oldFilterPtr = 0;
        _isRegistered = false;
    }

    /// <summary>
    /// Gets whether the OLE message filter is registered on the current thread.
    /// </summary>
    public static bool IsRegistered => _isRegistered;

    /// <summary>
    /// Gets whether the filter is currently in a long operation on this thread.
    /// </summary>
    public static bool IsInLongOperation => _isInLongOperation;

    /// <summary>
    /// Marks the beginning of a long-running COM operation (e.g., connection.Refresh()).
    /// While in a long operation, inbound COM callbacks are rejected with SERVERCALL_RETRYLATER
    /// instead of being dispatched, preventing CPU spin from re-entrant COM calls.
    /// </summary>
    public static void EnterLongOperation()
    {
        _messagePendingCount = 0;
        _handleInComingCallRejections = 0;
        _longOperationStartTimestamp = System.Diagnostics.Stopwatch.GetTimestamp();
        _isInLongOperation = true;
    }

    /// <summary>
    /// Marks the end of a long-running COM operation and returns diagnostic counters.
    /// </summary>
    /// <returns>A tuple of (messagePendingCalls, incomingCallRejections, elapsedMs).</returns>
    public static (long MessagePendingCalls, long IncomingCallRejections, double ElapsedMs) ExitLongOperation()
    {
        _isInLongOperation = false;
        var elapsed = System.Diagnostics.Stopwatch.GetElapsedTime(_longOperationStartTimestamp);
        return (_messagePendingCount, _handleInComingCallRejections, elapsed.TotalMilliseconds);
    }

    /// <summary>
    /// Associates a <see cref="CancellationToken"/> with the current STA thread's outgoing COM call.
    /// When the token is cancelled, <c>IMessageFilter.MessagePending</c> returns PENDINGMSG_CANCELCALL (0),
    /// causing the outgoing COM call (e.g., <c>connection.Refresh()</c>) to abort with
    /// RPC_E_CALL_CANCELLED rather than hanging until the caller's 30-minute timeout expires.
    /// </summary>
    /// <remarks>
    /// Must be called on the STA thread immediately before the COM call.
    /// Must be paired with <see cref="ClearPendingCancellationToken"/> in a finally block.
    /// </remarks>
    public static void SetPendingCancellationToken(CancellationToken token)
    {
        _pendingCancellationToken = token;
    }

    /// <summary>
    /// Clears the pending cancellation token after the COM call completes or is cancelled.
    /// Must be called in a finally block after every <see cref="SetPendingCancellationToken"/> call.
    /// </summary>
    public static void ClearPendingCancellationToken()
    {
        _pendingCancellationToken = default;
    }

    /// <summary>
    /// Handles incoming COM calls.
    /// During long operations, rejects with SERVERCALL_RETRYLATER to trigger the caller's
    /// RetryRejectedCall backoff mechanism, preventing CPU spin from re-entrant dispatch.
    /// </summary>
    int IOleMessageFilter.HandleInComingCall(int dwCallType, nint htaskCaller, int dwTickCount, nint lpInterfaceInfo)
    {
        if (_isInLongOperation)
        {
            // SERVERCALL_RETRYLATER (2) — reject but with proper COM retry protocol.
            // The COM runtime invokes the CALLER's IMessageFilter.RetryRejectedCall,
            // which implements backoff. This is fundamentally different from WAITNOPROCESS
            // rejection (which bypasses RetryRejectedCall and gives raw RPC_E_CALL_REJECTED).
            //
            // The callback is rejected BEFORE being dispatched to .NET, so no
            // EnsureScanDefinedEvents or IDispatch.TryGetTypeInfoCount runs.
            Interlocked.Increment(ref _handleInComingCallRejections);
            return 2; // SERVERCALL_RETRYLATER
        }

        // SERVERCALL_ISHANDLED (0) — accept the call (normal short operations)
        return 0;
    }

    /// <summary>
    /// Handles rejected COM calls from Excel.
    /// Implements automatic retry logic with exponential backoff for busy/unavailable conditions.
    /// Also retries SERVERCALL_REJECTED (modal dialogs like auth popups) with a longer window
    /// so the user has time to dismiss the dialog before the call is cancelled.
    /// </summary>
    /// <param name="htaskCallee">Handle to the task that rejected the call</param>
    /// <param name="dwTickCount">Number of milliseconds since rejection occurred</param>
    /// <param name="dwRejectType">Reason for rejection</param>
    /// <returns>
    /// 100+ = Retry after N milliseconds
    /// 0-99 = Cancel the call
    /// -1 = Cancel immediately
    /// </returns>
    int IOleMessageFilter.RetryRejectedCall(nint htaskCallee, int dwTickCount, int dwRejectType)
    {
        // dwRejectType values:
        // SERVERCALL_RETRYLATER (2) = Server is busy, try again later
        // SERVERCALL_REJECTED (1) = Server rejected the call (typically modal dialog showing)

        const int SERVERCALL_REJECTED = 1;
        const int SERVERCALL_RETRYLATER = 2;
        const int RETRY_TIMEOUT_MS = 30000;
        const int REJECTED_RETRY_TIMEOUT_MS = 120000; // 2 minutes for modal dialogs (auth popups, etc.)

        if (dwRejectType == SERVERCALL_RETRYLATER)
        {
            if (dwTickCount >= RETRY_TIMEOUT_MS)
            {
                return -1; // Cancel the call if timeout exceeded
            }

            // Exponential backoff based on elapsed time:
            // 0-1s:   100ms delays (quick retries for brief busy states)
            // 1-5s:   200ms delays
            // 5-15s:  500ms delays
            // 15-30s: 1000ms delays (Excel is seriously stuck)
            return dwTickCount switch
            {
                < 1000 => 100,
                < 5000 => 200,
                < 15000 => 500,
                _ => 1000
            };
        }

        if (dwRejectType == SERVERCALL_REJECTED)
        {
            // Modal dialog scenario (auth popups, privacy level dialogs, etc.)
            // Retry with longer backoff to give the user time to dismiss the dialog.
            // Without this, COM calls fail immediately when Excel shows any modal UI.
            if (dwTickCount >= REJECTED_RETRY_TIMEOUT_MS)
            {
                return -1; // Give up after 2 minutes — dialog was not dismissed
            }

            // Slower backoff since we're waiting for human interaction
            return dwTickCount switch
            {
                < 2000 => 500,
                < 10000 => 1000,
                _ => 2000
            };
        }

        return -1; // Cancel immediately for unknown rejection types
    }

    /// <summary>
    /// Handles pending message during a COM call.
    /// Context-dependent: during long operations, dispatches to HandleInComingCall (which rejects).
    /// During normal operations, queues messages without dispatching.
    ///
    /// The ExcelWriteGuard (integrated into Execute) suppresses ScreenUpdating to reduce
    /// the number of COM callbacks generated. For the remaining callbacks:
    /// - Long operations (refresh, Data Model): WAITDEFPROCESS → HandleInComingCall rejects
    /// - Normal operations: WAITNOPROCESS → messages queued until outbound call returns
    /// </summary>
    int IOleMessageFilter.MessagePending(nint htaskCallee, int dwTickCount, int dwPendingType)
    {
        Interlocked.Increment(ref _messagePendingCount);

        // If the operation's CancellationToken fired, abort the outgoing COM call immediately.
        // PENDINGMSG_CANCELCALL (0) — COM cancels the pending call; the outgoing call (e.g.,
        // connection.Refresh()) returns RPC_E_CALL_CANCELLED, unblocking the STA thread.
        // This prevents the thread from being permanently orphaned when the caller's timeout
        // fires while the thread is blocked inside a synchronous COM dispatch to Excel.
        if (_pendingCancellationToken.IsCancellationRequested)
        {
            return 0; // PENDINGMSG_CANCELCALL
        }

        if (_isInLongOperation)
        {
            // PENDINGMSG_WAITDEFPROCESS (2) — dispatch to HandleInComingCall.
            // During long operations, inbound COM callbacks are dispatched to HandleInComingCall
            // which rejects with SERVERCALL_RETRYLATER to trigger the caller's RetryRejectedCall
            // backoff, preventing CPU spin from re-entrant dispatch.
            return 2; // PENDINGMSG_WAITDEFPROCESS
        }

        // PENDINGMSG_WAITNOPROCESS (1) — queue inbound messages without dispatching.
        // For normal operations, don't dispatch inbound messages. Dispatching causes
        // re-entrant COM execution on the STA thread which hangs Data Model operations.
        // Messages are delivered after the outgoing call returns.
        return 1; // PENDINGMSG_WAITNOPROCESS
    }

    /// <summary>
    /// Registers or revokes a message filter for the current apartment.
    /// </summary>
    [LibraryImport("Ole32.dll")]
    private static partial int CoRegisterMessageFilter(
        nint lpMessageFilter,
        out nint lplpMessageFilter);
}


