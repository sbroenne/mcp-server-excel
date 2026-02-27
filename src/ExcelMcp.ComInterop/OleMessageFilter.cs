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
    /// Handles incoming COM calls. Not used for Excel automation scenarios.
    /// </summary>
    int IOleMessageFilter.HandleInComingCall(int dwCallType, nint htaskCaller, int dwTickCount, nint lpInterfaceInfo)
    {
        // SERVERCALL_ISHANDLED (0) - Accept the call
        return 0;
    }

    /// <summary>
    /// Handles rejected COM calls from Excel.
    /// Implements automatic retry logic with exponential backoff for busy/unavailable conditions.
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
        // SERVERCALL_REJECTED (1) = Server rejected the call

        const int SERVERCALL_RETRYLATER = 2;
        const int RETRY_TIMEOUT_MS = 30000;

        if (dwRejectType != SERVERCALL_RETRYLATER)
        {
            return -1; // Cancel immediately for non-retry scenarios
        }

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

    /// <summary>
    /// Handles pending message during a COM call.
    /// </summary>
    int IOleMessageFilter.MessagePending(nint htaskCallee, int dwTickCount, int dwPendingType)
    {
        // PENDINGMSG enum values (tagPENDINGMSG):
        //   PENDINGMSG_CANCELCALL    = 0  — cancel the outgoing call (dangerous, avoid)
        //   PENDINGMSG_WAITNOPROCESS = 1  — wait for return, do NOT dispatch the inbound message
        //   PENDINGMSG_WAITDEFPROCESS = 2 — wait and dispatch WM_PAINT / activation messages
        //
        // Return PENDINGMSG_WAITNOPROCESS (1) to block inbound COM dispatch.
        //
        // CRITICAL: Do NOT return PENDINGMSG_WAITDEFPROCESS (2) here.
        // That value dispatches inbound COM calls, causing two distinct failure modes:
        //
        // 1. STA deadlock (FormatConditions.Add scenario):
        //    a. STA thread calls FormatConditions.Add() → Excel starts processing
        //    b. Excel fires a callback (Calculate/SheetChange) back to this STA thread
        //    c. WAITDEFPROCESS dispatches it → callback tries to re-enter COM
        //    d. Excel waits for FormatConditions to finish → DEADLOCK
        //
        // 2. EnsureScanDefinedEvents CPU spin (PQ Data Model refresh scenario):
        //    a. STA thread calls dynamic connection.Refresh() → Excel processes Data Model
        //    b. MashupHost.exe fires hundreds of row-write callbacks per second to our STA
        //    c. WAITDEFPROCESS dispatches each one → .NET dynamic binder runs IDispatchMetaObject
        //       .BindGetMember for each callback, calling EnsureScanDefinedEvents
        //       → IDispatch.TryGetTypeInfoCount blocks on Excel (already busy) → tight spin
        //    d. Result: >97% CPU for the full duration of the Data Model write phase
        //
        // PENDINGMSG_WAITNOPROCESS (1) queues inbound messages without dispatching them.
        // They are delivered after the outgoing call returns. This is safe because:
        // - connection.Refresh(BackgroundQuery=false) is fully synchronous
        // - Excel does not need *our* STA to dispatch callbacks to make progress
        // - All pending callbacks are processed normally once Refresh() returns
        return 1; // PENDINGMSG_WAITNOPROCESS — queue inbound messages, do not dispatch
    }

    /// <summary>
    /// Registers or revokes a message filter for the current apartment.
    /// </summary>
    [LibraryImport("Ole32.dll")]
    private static partial int CoRegisterMessageFilter(
        nint lpMessageFilter,
        out nint lplpMessageFilter);
}


