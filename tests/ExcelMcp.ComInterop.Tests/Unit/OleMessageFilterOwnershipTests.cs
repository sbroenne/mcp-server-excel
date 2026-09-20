using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "ComInterop")]
public class OleMessageFilterOwnershipTests
{
    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(nint filter, out nint previous);

    [Fact]
    public void Register_ReleasesItsLocalInterfaceReference()
    {
        RunOnSta(() =>
        {
            OleMessageFilter.Register();
            nint registered = 0;
            try
            {
                Assert.Equal(0, CoRegisterMessageFilter(0, out registered));
                Assert.NotEqual(0, registered);
                Assert.Equal(1, ReferenceCount(registered));
            }
            finally
            {
                OleMessageFilter.Revoke();
                if (registered != 0) Marshal.Release(registered);
            }
        });
    }

    [Fact]
    public void Revoke_ReleasesReturnedFilterReference()
    {
        RunOnSta(() =>
        {
            OleMessageFilter.Register();
            nint registered = 0;
            try
            {
                Assert.Equal(0, CoRegisterMessageFilter(0, out registered));
                Assert.Equal(0, CoRegisterMessageFilter(registered, out var previous));
                Assert.Equal(0, previous);
                OleMessageFilter.Revoke();
                Assert.Equal(1, ReferenceCount(registered));
            }
            finally
            {
                OleMessageFilter.Revoke();
                if (registered != 0) Marshal.Release(registered);
            }
        });
    }

    [Fact]
    public void Revoke_RestoresPreviousFilterWithoutRetainingOwnedReferences()
    {
        RunOnSta(() =>
        {
            var wrappers = new StrategyBasedComWrappers();
            var original = wrappers.GetOrCreateComInterfaceForObject(new OleMessageFilter(), CreateComInterfaceFlags.None);
            nint restored = 0;
            try
            {
                Assert.Equal(0, CoRegisterMessageFilter(original, out var previous));
                Assert.Equal(0, previous);
                OleMessageFilter.Register();
                OleMessageFilter.Revoke();
                Assert.Equal(0, CoRegisterMessageFilter(0, out restored));
                Assert.Equal(original, restored);
                Marshal.Release(restored);
                restored = 0;
                Assert.Equal(1, ReferenceCount(original));
            }
            finally
            {
                OleMessageFilter.Revoke();
                _ = CoRegisterMessageFilter(0, out var remaining);
                if (remaining != 0) Marshal.Release(remaining);
                if (restored != 0) Marshal.Release(restored);
                Marshal.Release(original);
            }
        });
    }

    private static int ReferenceCount(nint pointer)
    {
        Marshal.AddRef(pointer);
        return Marshal.Release(pointer);
    }

    private static void RunOnSta(Action action)
    {
        Exception? failure = null;
        var thread = new Thread(() =>
        {
            try { action(); }
            catch (Exception ex) { failure = ex; }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        Assert.True(thread.Join(TimeSpan.FromSeconds(10)), "Message-filter test did not finish.");
        if (failure != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
    }
}
