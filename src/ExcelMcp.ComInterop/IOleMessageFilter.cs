using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>
/// COM interface for handling incoming and outgoing COM calls.
/// Used to intercept Excel busy/retry scenarios.
/// </summary>
[ComImport]
[Guid("00000016-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IOleMessageFilter
{
    [PreserveSig]
    int HandleInComingCall(
        int dwCallType,
        IntPtr htaskCaller,
        int dwTickCount,
        IntPtr lpInterfaceInfo);

    [PreserveSig]
    int RetryRejectedCall(
        IntPtr htaskCallee,
        int dwTickCount,
        int dwRejectType);

    [PreserveSig]
    int MessagePending(
        IntPtr htaskCallee,
        int dwTickCount,
        int dwPendingType);
}
