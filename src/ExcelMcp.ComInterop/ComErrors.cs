using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>Specific COM error codes used for optional member capability checks.</summary>
public static class ComErrors
{
    /// <summary>IDispatch could not find the requested member.</summary>
    public const int MemberNotFound = unchecked((int)0x80020003);

    /// <summary>IDispatch could not resolve the requested name.</summary>
    public const int UnknownName = unchecked((int)0x80020006);

    /// <summary>The member is not implemented by this server.</summary>
    public const int NotImplemented = unchecked((int)0x80004001);

    /// <summary>Whether an optional member is absent, rather than failing unexpectedly.</summary>
    public static bool IsMemberUnavailable(COMException exception) =>
        exception.HResult is MemberNotFound or UnknownName or NotImplemented;
}
