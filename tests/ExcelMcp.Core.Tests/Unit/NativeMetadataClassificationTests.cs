using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("RequiresExcel", "false")]
[Trait("Category", "Unit")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connection")]
[Trait("Speed", "Fast")]
public sealed class NativeMetadataClassificationTests
{
    private static readonly Func<Func<object?>, bool, DateTime?> ReadRefreshDate =
        typeof(ConnectionCommands).GetMethod("ReadOleDbRefreshDate", BindingFlags.Static | BindingFlags.NonPublic)!
            .CreateDelegate<Func<Func<object?>, bool, DateTime?>>();

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RefreshDate_SuccessfulTimestamp_IsRetainedRegardlessOfConnectionState(bool connected)
    {
        var expected = new DateTime(2026, 1, 2, 3, 4, 5);
        Assert.Equal(expected, ReadRefreshDate(() => expected, connected));
    }

    [Fact]
    public void RefreshDate_UnsetAndDisconnected_IsUnknown()
    {
        var failure = NativeFailure(unchecked((int)0x800A03EC));
        Assert.Null(ReadRefreshDate(() => throw failure, false));
    }

    [Theory]
    [InlineData(unchecked((int)0x800A03EC), true)]
    [InlineData(unchecked((int)0x80010108), false)]
    [InlineData(unchecked((int)0x800706BA), false)]
    [InlineData(ComErrors.MemberNotFound, false)]
    public void RefreshDate_NotCorroboratedUnset_PreservesOriginalFailure(int hresult, bool connected)
    {
        var failure = NativeFailure(hresult);
        Assert.Same(failure, Assert.Throws<COMException>(() => ReadRefreshDate(() => throw failure, connected)));
    }

    private static COMException NativeFailure(int hresult)
    {
#pragma warning disable CA2201 // Model exact native getter failures without Excel.
        return new COMException("Injected native timestamp failure.", hresult);
#pragma warning restore CA2201
    }
}
