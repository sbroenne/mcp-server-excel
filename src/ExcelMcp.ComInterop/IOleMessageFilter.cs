using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>
/// COM interface for handling incoming and outgoing COM calls.
/// Used to intercept Excel busy/retry scenarios.
/// </summary>
[GeneratedComInterface]
[Guid("00000016-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal partial interface IOleMessageFilter
{
    [PreserveSig]
    int HandleInComingCall(
        int dwCallType,
        nint htaskCaller,
        int dwTickCount,
        nint lpInterfaceInfo);

    [PreserveSig]
    int RetryRejectedCall(
        nint htaskCallee,
        int dwTickCount,
        int dwRejectType);

    [PreserveSig]
    int MessagePending(
        nint htaskCallee,
        int dwTickCount,
        int dwPendingType);
}


