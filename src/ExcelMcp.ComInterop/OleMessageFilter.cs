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

    /// <summary>
    /// Registers the OLE message filter for the current STA thread.
    /// Should be called once per STA thread before making COM calls.
    /// </summary>
    public static void Register()
    {
        var newFilter = new OleMessageFilter();
        nint newFilterPtr = s_comWrappers.GetOrCreateComInterfaceForObject(newFilter, CreateComInterfaceFlags.None);

        int result = CoRegisterMessageFilter(newFilterPtr, out _oldFilterPtr);
        if (result != 0)
        {
            throw new InvalidOperationException($"Failed to register OLE message filter. HRESULT: 0x{result:X8}");
        }
    }

    /// <summary>
    /// Revokes the OLE message filter and restores the previous filter.
    /// Should be called when STA thread is shutting down.
    /// </summary>
    public static void Revoke()
    {
        int result = CoRegisterMessageFilter(_oldFilterPtr, out _);
        if (result != 0)
        {
            throw new InvalidOperationException($"Failed to revoke OLE message filter. HRESULT: 0x{result:X8}");
        }
        _oldFilterPtr = 0;
    }

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
    /// Implements automatic retry logic for busy/unavailable conditions.
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

        // Early return pattern to reduce nesting
        const int SERVERCALL_RETRYLATER = 2;
        const int RETRY_TIMEOUT_MS = 30000;
        const int RETRY_DELAY_MS = 100;

        if (dwRejectType != SERVERCALL_RETRYLATER)
        {
            return -1; // Cancel immediately for non-retry scenarios
        }

        // Retry after 100ms for up to 30 seconds
        if (dwTickCount < RETRY_TIMEOUT_MS)
        {
            return RETRY_DELAY_MS; // Retry after 100ms
        }

        // Cancel the call if timeout exceeded
        return -1;
    }

    /// <summary>
    /// Handles pending message during a COM call.
    /// </summary>
    int IOleMessageFilter.MessagePending(nint htaskCallee, int dwTickCount, int dwPendingType)
    {
        // PENDINGMSG_WAITDEFPROCESS (2) - Continue waiting for the call to complete
        return 2;
    }

    /// <summary>
    /// Registers or revokes a message filter for the current apartment.
    /// </summary>
    [LibraryImport("Ole32.dll")]
    private static partial int CoRegisterMessageFilter(
        nint lpMessageFilter,
        out nint lplpMessageFilter);
}
